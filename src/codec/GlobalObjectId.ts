///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";
import { dateToFiletime } from "./PropertyValue.js";

/**
 * Encodes/decodes `CalendarEvent.icalUid` to/from the `GlobalObjectId` BLOB carried by `PidLidGlobalObjectId`
 * (`[MS-OXOCAL]` §2.2.1.27, confirmed field-by-field this session): `ByteArrayID(16, a fixed spec constant)
 * +YH(1)+YL(1)+M(1)+D(1)+CreationTime(8,FILETIME)+X(8,reserved,zero)+Size(4)+Data(variable, Size bytes)` - 40
 * fixed header bytes followed by `Data`.
 *
 * This server is always the one *generating* the `GlobalObjectId` in the first place (embedded in a meeting
 * invite it composes and sends, `RopSubmitMessageHandler.submitAppointment`) - a real client's own meeting
 * response always echoes back the identical `GlobalObjectId` its invite carried (per `[MS-OXOCAL]`'s own "this
 * property MUST NOT change" rule), so embedding `icalUid` directly in `Data` (as plain UTF-8 bytes) and reading
 * it back out is a direct, spec-legitimate round trip requiring no additional correlation table.
 *
 * **Pragmatic scope**: `YH`/`YL`/`M`/`D` (the `PidLidExceptionReplaceTime` fields, used only when a
 * `GlobalObjectId` identifies a single modified occurrence of a recurring series) are always encoded as `0`
 * ("not an exception") and not decoded - recurrence exceptions are a documented gap elsewhere in this pragmatic
 * subset (see `AppointmentRecurrence.ts`'s own doc comment), so there is nothing to source a real value from.
 * `PidLidCleanGlobalObjectId` (the sibling property real Outlook also sets, identical structure but with
 * `YH`/`YL`/`M`/`D` always zeroed) is therefore byte-identical to this codec's own output and needs no separate
 * implementation - the same value serves both properties.
 *
 * @author Jean-Philippe Steinmetz
 */

/** `ByteArrayID` (`[MS-OXOCAL]` §2.2.1.27): a fixed 16-byte constant identifying this BLOB as a `GlobalObjectId`
 * - the spec's own exact required byte sequence, not invented. */
const BYTE_ARRAY_ID = Buffer.from([0x04, 0x00, 0x00, 0x00, 0x82, 0x00, 0xe0, 0x00, 0x74, 0xc5, 0xb7, 0x10, 0x1a, 0x82, 0xe0, 0x08]);

/** Encodes `icalUid` into a `GlobalObjectId` BLOB, using `at` as the `CreationTime`. */
export function encodeGlobalObjectId(icalUid: string, at: Date): Buffer {
    const data = Buffer.from(icalUid, "utf-8");

    const writer = new BufferWriter();
    writer.writeBytes(BYTE_ARRAY_ID);
    writer.writeUInt8(0); // YH - not an exception, see class doc comment
    writer.writeUInt8(0); // YL
    writer.writeUInt8(0); // M
    writer.writeUInt8(0); // D
    writer.writeBigUInt64LE(dateToFiletime(at)); // CreationTime
    writer.writeBytes(Buffer.alloc(8)); // X - reserved, MUST be zero
    writer.writeUInt32LE(data.length); // Size
    writer.writeBytes(data);

    return writer.toBuffer();
}

/** Decodes a `GlobalObjectId` BLOB (read from `reader`'s current position) back into the `icalUid` it was
 * built from. Throws if `ByteArrayID` doesn't match the spec's own fixed constant - a real, spec-mandated
 * identity check, not an invented restriction. */
export function decodeGlobalObjectId(reader: BufferReader): string {
    const byteArrayId = reader.readBytes(16);
    if (!byteArrayId.equals(BYTE_ARRAY_ID)) {
        throw new Error("GlobalObjectId: ByteArrayID did not match the required [MS-OXOCAL] constant.");
    }
    reader.readUInt8(); // YH
    reader.readUInt8(); // YL
    reader.readUInt8(); // M
    reader.readUInt8(); // D
    reader.readBigUInt64LE(); // CreationTime - not needed to recover icalUid
    reader.readBytes(8); // X
    const size = reader.readUInt32LE();
    const data = reader.readBytes(size);

    return data.toString("utf-8");
}
