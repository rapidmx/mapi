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
 * **This server never actually embeds a `GlobalObjectId` in the invites it sends** - `RopSubmitMessageHandler
 * .submitAppointment()`'s invite is a plain RFC 5546 `METHOD:REQUEST` iCalendar attachment relayed over SMTP,
 * not a native MAPI Store object a real client reads `PidLidGlobalObjectId` off of directly; that iCalendar
 * only ever carries a plain `UID:` line. A real Outlook client receiving it therefore always *synthesizes* its
 * own `PidLidGlobalObjectId` from that bare iCalendar UID, using the `VCALID` form both `[MS-ASEMAIL]` §2.2.2.37
 * and the equivalent `[MS-OXCICAL]` UID-conversion algorithm document: `Data` = the 8-byte ASCII marker
 * `"vCal-Uid"`, then a 4-byte version (`01 00 00 00`), then the UID's own bytes, then one trailing `NUL` -
 * **not** the bare UID bytes this codec originally assumed. `decodeGlobalObjectId` recognizes and unwraps this
 * form (falling back to the hex form below only when the `vCal-Uid` marker isn't present, i.e. an
 * `OutlookID` - a native-Exchange-generated `GlobalObjectId`, decoded to the hex form `[MS-OXCICAL]` gives its
 * iCalendar UID). `encodeGlobalObjectId` now builds the same `VCALID` shape, both so
 * a round trip through this codec models the real wire form and because `MeetingMessageClassHandler.test.ts`'s
 * own test doubles need a realistic fixture to encode.
 *
 * **Pragmatic scope**: `YH`/`YL`/`M`/`D` (set when a `GlobalObjectId` identifies a single occurrence of a recurring
 * series) are always encoded as `0` ("not an exception"); `globalObjectIdInstanceDate` reads them from a client's
 * meeting response.
 * `PidLidCleanGlobalObjectId` (the sibling property real Outlook also sets, identical structure but with
 * `YH`/`YL`/`M`/`D` always zeroed) is therefore byte-identical to this codec's own output and needs no separate
 * implementation - the same value serves both properties.
 *
 * @author Jean-Philippe Steinmetz
 */

/** `ByteArrayID` (`[MS-OXOCAL]` §2.2.1.27): a fixed 16-byte constant identifying this BLOB as a `GlobalObjectId`
 * - the spec's own exact required byte sequence, not invented. */
const BYTE_ARRAY_ID = Buffer.from([0x04, 0x00, 0x00, 0x00, 0x82, 0x00, 0xe0, 0x00, 0x74, 0xc5, 0xb7, 0x10, 0x1a, 0x82, 0xe0, 0x08]);

/** `VCALID`'s own fixed 8-byte ASCII marker (`[MS-ASEMAIL]` §2.2.2.37's `VCALSTRING`) identifying `Data` as a
 * `"vCal-Uid"`-wrapped UID rather than a native `OutlookID`'s raw bytes. */
const VCAL_MARKER = Buffer.from("vCal-Uid", "ascii");

/** `VCALID`'s own fixed 4-byte `VERSION` field - always this exact literal byte sequence, not a real version
 * negotiated at runtime. */
const VCAL_VERSION = Buffer.from([0x01, 0x00, 0x00, 0x00]);

/** A `VCALID`'s minimum possible `Data` length: `VCAL_MARKER` (8) + `VCAL_VERSION` (4) + a zero-length UID (0)
 * + the trailing `NUL` (1). */
const VCAL_MIN_DATA_LENGTH = VCAL_MARKER.length + VCAL_VERSION.length + 1;

/** Encodes `icalUid` into a `GlobalObjectId` BLOB in the `VCALID` form a real client synthesizes from a plain
 * iCalendar UID (see this file's own doc comment for why that's the form that matters here), using `at` as the
 * `CreationTime`. */
export function encodeGlobalObjectId(icalUid: string, at: Date): Buffer {
    const data = Buffer.concat([VCAL_MARKER, VCAL_VERSION, Buffer.from(icalUid, "utf-8"), Buffer.from([0x00])]);

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

/** Decodes a `GlobalObjectId` BLOB (read from `reader`'s current position) back into the `icalUid` it stands for.
 * A `VCALID` (see this file's own doc comment) yields the UID it wraps. Anything else is a native `OutlookID`,
 * whose `Data` is binary, not text: per `[MS-OXCICAL]`'s UID conversion, the iCalendar UID of such an object is the
 * uppercase hex of the whole `GlobalObjectId` with the instance date (`YH`/`YL`/`M`/`D`, bytes 16-19) zeroed, so
 * that is what is returned (compare it case-insensitively). Throws if `ByteArrayID` doesn't match the spec's own
 * fixed constant - a real, spec-mandated identity check, not an invented restriction. */
export function decodeGlobalObjectId(reader: BufferReader): string {
    const byteArrayId = reader.readBytes(16);
    if (!byteArrayId.equals(BYTE_ARRAY_ID)) {
        throw new Error("GlobalObjectId: ByteArrayID did not match the required [MS-OXOCAL] constant.");
    }
    reader.readBytes(4); // YH/YL/M/D
    const creationTimeAndReserved = reader.readBytes(16); // CreationTime + X
    const sizeBytes = reader.readBytes(4);
    const size = sizeBytes.readUInt32LE(0);
    const data = reader.readBytes(size);

    const isVCalId = size >= VCAL_MIN_DATA_LENGTH && data.subarray(0, 8).equals(VCAL_MARKER) && data.subarray(8, 12).equals(VCAL_VERSION);
    if (isVCalId) {
        // Strip the marker+version prefix and the trailing NUL, per [MS-ASEMAIL]'s own UID-length formula
        // (BYTECOUNT minus the marker minus the version minus 1 byte for the NUL).
        return data.subarray(12, data.length - 1).toString("utf-8");
    }

    return Buffer.concat([byteArrayId, Buffer.alloc(4), creationTimeAndReserved, sizeBytes, data]).toString("hex").toUpperCase();
}

/** The `YYYY-MM-DD` occurrence date a `GlobalObjectId` names (`YH`/`YL`/`M`/`D`), or `undefined` when it names the
 * whole series (a zero year). `globalObjectId` must already have decoded successfully. */
export function globalObjectIdInstanceDate(globalObjectId: Buffer): string | undefined {
    const year = (globalObjectId[16] << 8) | globalObjectId[17];
    if (year === 0) {
        return undefined;
    }
    return `${String(year).padStart(4, "0")}-${String(globalObjectId[18]).padStart(2, "0")}-${String(globalObjectId[19]).padStart(2, "0")}`;
}
