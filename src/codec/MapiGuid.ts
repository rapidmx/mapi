///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";

/**
 * Encodes/decodes the 16-byte binary GUID layout `[MS-DTYP]`'s `GUID` structure defines and every MAPI wire
 * structure that embeds a GUID (EntryIDs, `PidTagStoreRecordKey`, `PtypGuid` property values, property set
 * IDs) uses: `Data1`/`Data2`/`Data3` in little-endian order, followed by `Data4`'s 8 bytes unchanged (the same
 * order they appear in a standard textual GUID's last two hyphen-separated groups). This is distinct from a
 * plain RFC 4122 UUID's all-big-endian textual byte order, so a generic UUID library's byte form can't be
 * reused directly - hand-built here, same "no library, build the wire format precisely" precedent as the
 * WBXML codec and the Autodiscover XML helpers.
 *
 * @author Jean-Philippe Steinmetz
 */
const GUID_STRING_PATTERN = /^([0-9a-f]{8})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{12})$/i;

/** Encodes a standard hyphenated GUID/UUID string into its 16-byte MS wire form. */
export function encodeGuid(guid: string): Buffer {
    const match = GUID_STRING_PATTERN.exec(guid);
    if (!match) {
        throw new Error(`Invalid GUID string: ${guid}`);
    }
    const [, data1, data2, data3, data4a, data4b] = match;

    const writer = new BufferWriter();
    writer.writeUInt32LE(parseInt(data1, 16));
    writer.writeUInt16LE(parseInt(data2, 16));
    writer.writeUInt16LE(parseInt(data3, 16));
    writer.writeBytes(Buffer.from(data4a + data4b, "hex"));
    return writer.toBuffer();
}

/** Decodes a 16-byte MS wire-form GUID (read from `reader`'s current position) into a standard hyphenated
 * GUID/UUID string. */
export function decodeGuid(reader: BufferReader): string {
    const data1 = reader.readUInt32LE().toString(16).padStart(8, "0");
    const data2 = reader.readUInt16LE().toString(16).padStart(4, "0");
    const data3 = reader.readUInt16LE().toString(16).padStart(4, "0");
    const data4 = reader.readBytes(8).toString("hex");
    return `${data1}-${data2}-${data3}-${data4.slice(0, 4)}-${data4.slice(4)}`;
}
