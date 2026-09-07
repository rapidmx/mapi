///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// GlobalObjectId is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// test/mapi/codec/AppointmentRecurrence.test.ts.
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { decodeGlobalObjectId, encodeGlobalObjectId } from "../../src/codec/GlobalObjectId.js";

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

    it("Encodes Size as the exact byte length of the UTF-8-encoded icalUid.", () => {
        const icalUid = "a-longer-unique-id-1234567890@example.com";
        const encoded = encodeGlobalObjectId(icalUid, new Date());
        expect(encoded.readUInt32LE(36)).toBe(Buffer.byteLength(icalUid, "utf-8"));
        expect(encoded.length).toBe(40 + Buffer.byteLength(icalUid, "utf-8"));
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
