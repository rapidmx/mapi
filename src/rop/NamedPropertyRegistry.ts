///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { MapiSessionContext } from "../MapiSessionManager.js";

/** A `PropertyName` (`[MS-OXCDATA]` §2.6.1): identifies a named property by a property-set `guid` plus either a
 * numeric `lid` (`kind: "lid"`) or a string `name` (`kind: "name"`) - never both. `guid` is the textual form
 * `MapiGuid.ts`'s `encodeGuid`/`decodeGuid` already produce/consume (the wire `PropertyName` structure's own
 * `GUID` field is a `FlatUID` - confirmed, per its own spec page, to be byte-identical to the same little-endian
 * `Data1`/`Data2`/`Data3`+as-is-`Data4` layout `MapiGuid.ts` already implements, so no second GUID codec is
 * needed here). */
export interface PropertyName {
    guid: string;
    kind: "lid" | "name";
    lid?: number;
    name?: string;
}

/** The first numeric property ID this pragmatic subset ever assigns to a named property - real Exchange also
 * reserves the `0x0000`-`0x7FFF` range for well-known/`PidTag*` properties, only ever assigning named
 * properties IDs at `0x8000` and above (`[MS-OXCPRPT]`'s own `RopGetPropertyIdsFromNames` processing rules). */
const FIRST_NAMED_PROPERTY_ID = 0x8000;

/** A `PropertyName`'s registry key - a JSON string rather than a delimiter-joined one (e.g. `"<guid>:<lid>"`)
 * because a string-`Kind` named property's own `name` is an arbitrary client-supplied string that could
 * legitimately contain any delimiter this codec might otherwise pick (some real named properties do use
 * URN-shaped names). `JSON.stringify` on a small, fixed-shape object has no such ambiguity. */
function keyFor(propertyName: PropertyName): string {
    return propertyName.kind === "lid"
        ? JSON.stringify({ guid: propertyName.guid.toLowerCase(), kind: "lid", lid: propertyName.lid })
        : JSON.stringify({ guid: propertyName.guid.toLowerCase(), kind: "name", name: propertyName.name });
}

/** The highest numeric property ID a `PropertyTag`'s 16-bit `PropertyId` field can carry at all - not a
 * pragmatic-subset choice, an absolute wire-format ceiling (`writeUInt16LE` cannot encode anything past this). */
const LAST_NAMED_PROPERTY_ID = 0xffff;

/**
 * Returns `propertyName`'s existing numeric property ID if an earlier `RopGetPropertyIdsFromNames` call in
 * this session already assigned one, otherwise assigns and remembers the next free ID starting at
 * `FIRST_NAMED_PROPERTY_ID`. This pragmatic subset always behaves as if the request's own `Flags` field
 * requested "assign a new ID if unmapped" (`0x02`) - real Exchange's alternative (`0x00`, "only return
 * already-mapped IDs") exists to let a client probe without committing a mailbox-wide registration; since this
 * registry is already only ever session-scoped (not a real persisted per-mailbox mapping table), there is no
 * meaningful difference between "probe" and "assign" here.
 *
 * Backed by `session.namedProperties`/`namedPropertyIds` (a forward and reverse map, kept in sync) and
 * `session.nextNamedPropertyId`, an O(1) lookup/assignment pair rather than a linear scan plus a
 * `Math.max(...spread)` over `Object.values(session.namedProperties)` - both real costs at scale: the spread
 * form risks a `RangeError: Maximum call stack size exceeded` once a session has registered enough distinct
 * names to exceed V8's function-argument-count limit, and the registry only ever grows for a session's
 * lifetime, so a linear scan repeated once per property per row is quadratic over a session that resolves many
 * named properties.
 *
 * Once `nextNamedPropertyId` would exceed `LAST_NAMED_PROPERTY_ID` (32,768 distinct names already registered
 * this session - the entire `0x8000`-`0xFFFF` numeric ID space this pragmatic subset has to hand out), a *new*
 * name can no longer be assigned a real ID; this returns `0x0000` for it instead of throwing, the exact value
 * `[MS-OXCPRPT]`'s own `RopGetPropertyIdsFromNames` processing rules already use for "this `PropertyName`
 * could not be resolved" (the same value this pragmatic subset already produces for a `Kind = 0xFF` entry) -
 * not a new failure mode, just the same one applied to a different unmappable case. A name already registered
 * before the registry filled up keeps returning its real, previously-assigned ID.
 */
export function assignOrGetNamedPropertyId(session: MapiSessionContext, propertyName: PropertyName): number {
    const key = keyFor(propertyName);
    const existing = session.namedProperties[key];
    if (existing !== undefined) {
        return existing;
    }
    if (session.nextNamedPropertyId > LAST_NAMED_PROPERTY_ID) {
        return 0x0000;
    }
    const id = session.nextNamedPropertyId++;
    session.namedProperties[key] = id;
    session.namedPropertyIds[id] = key;
    return id;
}

/** The reverse lookup `RopSetProperties`/`RopGetPropertiesSpecific`/`RopQueryRows` use to recognize an incoming
 * property ID `>= 0x8000` as one of this session's own mapped named properties - an O(1) lookup against
 * `session.namedPropertyIds` (kept in sync by `assignOrGetNamedPropertyId`) rather than a linear scan of
 * `session.namedProperties` with a `JSON.parse` per candidate - see that function's own doc comment for why
 * this matters at scale, particularly since this is called once per named-property column per row from
 * `RopQueryRows`. */
export function resolveNamedProperty(session: MapiSessionContext, propertyId: number): PropertyName | undefined {
    const key = session.namedPropertyIds[propertyId];
    return key !== undefined ? (JSON.parse(key) as PropertyName) : undefined;
}
