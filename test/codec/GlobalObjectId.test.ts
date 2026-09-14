///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// GlobalObjectId is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// test/mapi/codec/AppointmentRecurrence.test.ts.
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { decodeGlobalObjectId, encodeGlobalObjectId, globalObjectIdInstanceDate } from "../../src/codec/GlobalObjectId.js";

describe("GlobalObjectId Tests", () => {
    it("Round-trips an icalUid.", () => {
        const encoded = encodeGlobalObjectId("event-123@example.com", new Date("2026-09-07T14:00:00.000Z"));
        expect(decodeGlobalObjectId(new BufferReader(encoded))).toBe("event-123@example.com");
    });

    it("Writes the exact 16-byte ByteArrayID constant required by MS-OXOCAL.", () => {
        const encoded = encodeGlobalObjectId("uid", new Date());
        expect(encoded.subarray(0, 16)).toEqual(
            Buffer.from([0x04, 0x00, 0x00, 0x00, 0x82, 0x00, 0xe0, 0x00, 0x74, 0xc5, 0xb7, 0x10, 0x1a, 0x82, 0xe0, 0x08]),
        );
    });

    it("Writes zero for YH/YL/M/D (no recurrence-exception support) and all-zero X.", () => {
        const encoded = encodeGlobalObjectId("uid", new Date());
        expect(encoded.subarray(16, 20)).toEqual(Buffer.alloc(4)); // YH/YL/M/D
        expect(encoded.subarray(28, 36)).toEqual(Buffer.alloc(8)); // X
    });

    it("Encodes Size as the VCALID marker+version+UID+NUL length, not the bare UID's own byte length.", () => {
        const icalUid = "a-longer-unique-id-1234567890@example.com";
        const encoded = encodeGlobalObjectId(icalUid, new Date());
        // VCALID = "vCal-Uid" (8) + version (4) + UID + NUL (1).
        const expectedDataLength = 8 + 4 + Buffer.byteLength(icalUid, "utf-8") + 1;
        expect(encoded.readUInt32LE(36)).toBe(expectedDataLength);
        expect(encoded.length).toBe(40 + expectedDataLength);
    });

    it("Encodes Data in the real VCALID wire form: the vCal-Uid marker, then a 01 00 00 00 version, then the UID, then a trailing NUL.", () => {
        const encoded = encodeGlobalObjectId("uid-123", new Date());
        const data = encoded.subarray(40);
        expect(data.subarray(0, 8).toString("ascii")).toBe("vCal-Uid");
        expect(data.subarray(8, 12)).toEqual(Buffer.from([0x01, 0x00, 0x00, 0x00]));
        expect(data.subarray(12, data.length - 1).toString("utf-8")).toBe("uid-123");
        expect(data[data.length - 1]).toBe(0x00);
    });

    it("Decodes a native OutlookID (no vCal-Uid marker) as the uppercase hex of the whole id with the instance date zeroed, not as text.", () => {
        const writer = new BufferWriter();
        const byteArrayId = Buffer.from([0x04, 0x00, 0x00, 0x00, 0x82, 0x00, 0xe0, 0x00, 0x74, 0xc5, 0xb7, 0x10, 0x1a, 0x82, 0xe0, 0x08]);
        writer.writeBytes(byteArrayId);
        writer.writeUInt8(0x07);
        writer.writeUInt8(0xea);
        writer.writeUInt8(10);
        writer.writeUInt8(8);
        writer.writeBigUInt64LE(0n);
        writer.writeBytes(Buffer.alloc(8));
        const rawData = Buffer.from([0xff, 0xfe, 0x00, 0x81]); // binary, not valid UTF-8
        writer.writeUInt32LE(rawData.length);
        writer.writeBytes(rawData);
        const blob = writer.toBuffer();
        const expected = Buffer.from(blob);
        expected.fill(0, 16, 20);

        expect(decodeGlobalObjectId(new BufferReader(blob))).toBe(expected.toString("hex").toUpperCase());
        expect(decodeGlobalObjectId(new BufferReader(blob))).toMatch(/^040000008200E00074C5B7101A82E008000000000/);
        expect(globalObjectIdInstanceDate(blob)).toBe("2026-10-08");
        expect(globalObjectIdInstanceDate(expected)).toBeUndefined();
    });

    it("Round-trips a non-ASCII icalUid correctly via UTF-8 byte length (not string length).", () => {
        const icalUid = "événement-42@example.com";
        const encoded = encodeGlobalObjectId(icalUid, new Date());
        expect(decodeGlobalObjectId(new BufferReader(encoded))).toBe(icalUid);
    });

    it("Throws decoding a blob whose ByteArrayID doesn't match the required MS-OXOCAL constant.", () => {
        const encoded = encodeGlobalObjectId("uid", new Date());
        encoded.writeUInt8(0xff, 0); // corrupt the first ByteArrayID byte
        expect(() => decodeGlobalObjectId(new BufferReader(encoded))).toThrow(/ByteArrayID/);
    });
});
