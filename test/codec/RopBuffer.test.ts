///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// RopBuffer is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// the other test/mapi/codec/*.test.ts files.
import { decodeRopBuffer, encodeRopBuffer, type RopBuffer } from "../../src/codec/RopBuffer.js";

describe("RopBuffer Tests", () => {
    it("Encodes a hand-verifiable buffer to the exact expected byte sequence.", () => {
        // ropsList = [0xAA, 0xBB, 0xCC] (3 bytes) -> RopSize = 2 (itself) + 3 = 5 (LE: 05 00).
        // handleTable = [1, 2] -> 4 bytes each, LE: 01 00 00 00, 02 00 00 00.
        const buf: RopBuffer = { ropsList: Buffer.from([0xaa, 0xbb, 0xcc]), handleTable: [1, 2] };
        const encoded = encodeRopBuffer(buf);
        expect(encoded.toString("hex")).toBe("0500aabbcc0100000002000000");
    });

    it("Round-trips an empty ropsList and empty handle table.", () => {
        const buf: RopBuffer = { ropsList: Buffer.alloc(0), handleTable: [] };
        const decoded = decodeRopBuffer(encodeRopBuffer(buf));
        expect(decoded.ropsList.length).toBe(0);
        expect(decoded.handleTable).toEqual([]);
    });

    it("Round-trips a non-trivial ropsList and handle table.", () => {
        const buf: RopBuffer = {
            ropsList: Buffer.from([0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]),
            handleTable: [0, 1, 0xffffffff, 42],
        };
        const decoded = decodeRopBuffer(encodeRopBuffer(buf));
        expect(Buffer.compare(decoded.ropsList, buf.ropsList)).toBe(0);
        expect(decoded.handleTable).toEqual(buf.handleTable);
    });

    it("Decodes the exact expected byte sequence back to the original buffer.", () => {
        const bytes = Buffer.from("0500aabbcc0100000002000000", "hex");
        const decoded = decodeRopBuffer(bytes);
        expect(decoded.ropsList.toString("hex")).toBe("aabbcc");
        expect(decoded.handleTable).toEqual([1, 2]);
    });

    it("Throws on a malformed RopSize smaller than its own 2-byte field, instead of silently rewinding the cursor into the handle table.", () => {
        // RopSize=0 would previously compute a ropsList length of -2, which Buffer.subarray silently clamped
        // to empty while also rewinding the cursor's offset backwards - misparsing whatever followed as the
        // handle table instead of failing loudly on this malformed input.
        const bytes = Buffer.from("0000aabbcc0100000002000000", "hex");
        expect(() => decodeRopBuffer(bytes)).toThrow(RangeError);
    });
});
