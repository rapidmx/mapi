///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { decodeGuid } from "../codec/MapiGuid.js";
import { assignOrGetNamedPropertyId, type PropertyName } from "./NamedPropertyRegistry.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_PROPERTY_IDS_FROM_NAMES = 0x56;

/** `PropertyName.Kind` (`[MS-OXCDATA]` §2.6.1) values. */
const KIND_LID = 0x00;
const KIND_NAME = 0x01;
// KIND_NONE = 0xFF ("the property does not have an associated PropertyName field") is a real, spec-defined
// value but never resolvable to a real named property - handled by readPropertyName()'s default/`undefined`
// return rather than its own named constant.

/** The trailing UTF-16LE null-terminator code unit `PropertyName.Name` (Kind=0x01) is documented to always
 * carry - stripped by `charCodeAt`/numeric comparison rather than a string literal containing the null
 * character itself, to avoid embedding a raw control character in source text. */
const NULL_CODE_UNIT = 0;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle doesn't exist" - the
 * same constant `RopGetPropertiesSpecificHandler` uses for its own analogous check. Real Exchange resolves
 * named properties per-logon regardless of which specific object handle a client happens to call this ROP
 * against (Store/Folder/Message all share one mailbox-wide mapping table) - this pragmatic subset only
 * requires *some* real handle to exist, not a particular type, matching that "any object will do" semantics. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** Decodes one `PropertyName` structure (`[MS-OXCDATA]` §2.6.1, confirmed field order this session: `Kind`
 * **then** `GUID` - not the reverse). Returns `undefined` for `Kind = 0xFF` ("no associated PropertyName"),
 * which cannot resolve to any real named property. */
function readPropertyName(reader: BufferReader): PropertyName | undefined {
    const kind = reader.readUInt8();
    const guid = decodeGuid(reader);
    if (kind === KIND_LID) {
        return { guid, kind: "lid", lid: reader.readUInt32LE() };
    }
    if (kind === KIND_NAME) {
        const nameSize = reader.readUInt8();
        const nameBytes = reader.readBytes(nameSize);
        // The spec text describing NameSize ("the number of bytes in the Name string that follows it") doesn't
        // unambiguously say whether that count includes the field's own trailing 2-zero-byte terminator - this
        // pragmatic subset treats it as inclusive (the common real-world convention for this style of
        // length-prefixed MAPI string) and strips one trailing null code unit if present. Never exercised by
        // this pragmatic subset's own Calendar property table (every property it needs is Kind=LID), so this
        // is defensive completeness for an arbitrary client request, not a property on the critical path.
        let name = nameBytes.toString("utf16le");
        if (name.length > 0 && name.charCodeAt(name.length - 1) === NULL_CODE_UNIT) {
            name = name.slice(0, -1);
        }
        return { guid, kind: "name", name };
    }
    return undefined;
}

/**
 * `RopGetPropertyIdsFromNames` (`[MS-OXCPRPT]`/`[MS-OXCROPS]` §2.2.8.1): resolves a list of named properties
 * (`PropertyName` structures - a property-set `GUID` plus a numeric `LID` or string `Name`) into numeric
 * property IDs a subsequent `RopSetProperties`/`RopGetPropertiesSpecific`/`RopQueryRows` call can use directly.
 * This is the mechanism almost every Calendar-specific property (`PidLidAppointmentStartWhole`,
 * `PidLidBusyStatus`, `PidLidAppointmentRecur`, ...) requires, since none of them have a fixed numeric
 * `PidTag*`-style ID the way `PidTagSubject` does - see `NamedPropertyRegistry.ts`'s own doc comment for the
 * session-scoped registry this delegates to.
 *
 * An unresolvable `PropertyName` (`Kind = 0xFF`) maps to `0x0000` in the response, per spec - this pragmatic
 * subset has no other unmappable case (no permission/quota/hard-limit checks), so every `Kind = 0x00`/`0x01`
 * entry always succeeds. `Flags` is decoded to advance past it correctly but not honored - see
 * `assignOrGetNamedPropertyId`'s own doc comment for why "probe" and "assign" collapse to the same behavior
 * for a purely session-scoped registry.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopGetPropertyIdsFromNamesHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_PROPERTY_IDS_FROM_NAMES;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // Flags - see class doc comment
        const propertyNameCount: number = reader.readUInt16LE();
        const propertyNames: (PropertyName | undefined)[] = [];
        for (let i = 0; i < propertyNameCount; i++) {
            propertyNames.push(readPropertyName(reader));
        }

        if (!context.session.handles[inputHandleIndex]) {
            writer.writeUInt8(ROP_ID_GET_PROPERTY_IDS_FROM_NAMES);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const propertyIds = propertyNames.map((name) => (name ? assignOrGetNamedPropertyId(context.session, name) : 0x0000));

        writer.writeUInt8(ROP_ID_GET_PROPERTY_IDS_FROM_NAMES);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(propertyIds.length);
        for (const id of propertyIds) {
            writer.writeUInt16LE(id);
        }
    }
}
