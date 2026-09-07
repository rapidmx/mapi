///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// MapiGuid is pure binary-format logic with no DI/DB dependency, so it's tested directly here rather than
// only indirectly through a real-server HTTP round trip - the same precedent test/eas/codec/WbxmlCodec.test.ts
// sets for the WBXML codec.
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { decodeGuid, encodeGuid } from "../../src/codec/MapiGuid.js";

describe("MapiGuid Tests", () => {
    it("Encodes a hand-verifiable GUID to the exact expected byte sequence.", () => {
        // 01020304-0506-0708-090a-0b0c0d0e0f10 -> Data1=0x01020304 (LE: 04 03 02 01),
        // Data2=0x0506 (LE: 06 05), Data3=0x0708 (LE: 08 07), Data4 unchanged (09 0a 0b 0c 0d 0e 0f 10).
        const encoded = encodeGuid("01020304-0506-0708-090a-0b0c0d0e0f10");
        expect(encoded.toString("hex")).toBe("0403020106050807090a0b0c0d0e0f10");
    });

    it("Decodes the exact expected byte sequence back to the original GUID string.", () => {
        const bytes = Buffer.from("0403020106050807090a0b0c0d0e0f10", "hex");
        const decoded = decodeGuid(new BufferReader(bytes));
        expect(decoded).toBe("01020304-0506-0708-090a-0b0c0d0e0f10");
    });

    it("Round-trips a variety of GUID strings, case-insensitively on input.", () => {
        const guids = [
            "00000000-0000-0000-0000-000000000000",
            "ffffffff-ffff-ffff-ffff-ffffffffffff",
            "00062002-0000-0000-C000-000000000046",
            "12345678-9ABC-DEF0-1234-56789ABCDEF0",
        ];
        for (const guid of guids) {
            const encoded = encodeGuid(guid);
            expect(encoded.length).toBe(16);
            const decoded = decodeGuid(new BufferReader(encoded));
            expect(decoded).toBe(guid.toLowerCase());
        }
    });

    it("Decodes starting from a non-zero reader offset without disturbing prior bytes.", () => {
        const prefix = Buffer.from([0xaa, 0xbb]);
        const guidBytes = encodeGuid("12345678-9abc-def0-1234-56789abcdef0");
        const reader = new BufferReader(Buffer.concat([prefix, guidBytes]), 2);
        expect(decodeGuid(reader)).toBe("12345678-9abc-def0-1234-56789abcdef0");
        expect(reader.position).toBe(18);
    });

    it("Throws for a malformed GUID string.", () => {
        expect(() => encodeGuid("not-a-guid")).toThrow(/Invalid GUID string/);
    });
});
