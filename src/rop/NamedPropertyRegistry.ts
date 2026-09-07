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

/**
 * Returns `propertyName`'s existing numeric property ID if an earlier `RopGetPropertyIdsFromNames` call in
 * this session already assigned one, otherwise assigns and remembers the next free ID starting at
 * `FIRST_NAMED_PROPERTY_ID` - the exact `FolderTarget.assignOrGetFid`/`MessageTarget.assignOrGetMid` linear-
 * registry pattern already used twice in this codebase, adapted for named properties. This pragmatic subset
 * always behaves as if the request's own `Flags` field requested "assign a new ID if unmapped" (`0x02`) -
 * real Exchange's alternative (`0x00`, "only return already-mapped IDs") exists to let a client probe without
 * committing a mailbox-wide registration; since this registry is already only ever session-scoped (not a real
 * persisted per-mailbox mapping table), there is no meaningful difference between "probe" and "assign" here.
 */
export function assignOrGetNamedPropertyId(session: MapiSessionContext, propertyName: PropertyName): number {
    const key = keyFor(propertyName);
    const existing = session.namedProperties[key];
    if (existing !== undefined) {
        return existing;
    }
    const nextId = Math.max(FIRST_NAMED_PROPERTY_ID - 1, ...Object.values(session.namedProperties)) + 1;
    session.namedProperties[key] = nextId;
    return nextId;
}

/** The reverse lookup `RopSetProperties`/`RopGetPropertiesSpecific`/`RopQueryRows` use to recognize an incoming
 * property ID `>= 0x8000` as one of this session's own mapped named properties. A linear scan, matching this
 * codebase's own established precedent (`FolderTarget.assignOrGetFid`'s doc comment) that a single session's
 * named-property count is never large enough for this to be a real cost. */
export function resolveNamedProperty(session: MapiSessionContext, propertyId: number): PropertyName | undefined {
    for (const [key, id] of Object.entries(session.namedProperties)) {
        if (id === propertyId) {
            return JSON.parse(key) as PropertyName;
        }
    }
    return undefined;
}
