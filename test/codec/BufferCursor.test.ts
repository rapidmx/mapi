///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// BufferReader/BufferWriter are pure binary-format logic with no DI/DB dependency, tested directly here -
// same precedent as test/eas/codec/WbxmlCodec.test.ts. MapiGuid.test.ts/PropertyValue.test.ts already
// exercise the primitive read/write methods incidentally; this file covers BufferReader's own
// position/remaining/hasMore bookkeeping and the null-terminated string readers directly, since those aren't
// otherwise exercised until the RopBuffer framing codec (a later build step) starts consuming them.
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";

describe("BufferCursor Tests", () => {
    describe("BufferReader", () => {
        it("Tracks position/remaining/hasMore as bytes are consumed.", () => {
            const reader = new BufferReader(Buffer.from([0x01, 0x02, 0x03, 0x04]));
            expect(reader.position).toBe(0);
            expect(reader.remaining).toBe(4);
            expect(reader.hasMore()).toBe(true);

            reader.readUInt16LE();
            expect(reader.position).toBe(2);
            expect(reader.remaining).toBe(2);
            expect(reader.hasMore()).toBe(true);

            reader.readUInt16LE();
            expect(reader.position).toBe(4);
            expect(reader.remaining).toBe(0);
            expect(reader.hasMore()).toBe(false);
        });

        it("Accepts a non-zero starting offset.", () => {
            const reader = new BufferReader(Buffer.from([0xaa, 0xbb, 0x01, 0x00]), 2);
            expect(reader.position).toBe(2);
            expect(reader.readUInt16LE()).toBe(1);
        });

        it("readBytes() returns a subarray of the requested length and advances position.", () => {
            const reader = new BufferReader(Buffer.from([1, 2, 3, 4, 5]));
            const slice = reader.readBytes(3);
            expect([...slice]).toEqual([1, 2, 3]);
            expect(reader.position).toBe(3);
        });

        it("readNullTerminatedUtf16LE() stops at the terminator and advances past it.", () => {
            const buf = Buffer.concat([Buffer.from("hi", "utf16le"), Buffer.from([0x00, 0x00]), Buffer.from([0xff])]);
            const reader = new BufferReader(buf);
            expect(reader.readNullTerminatedUtf16LE()).toBe("hi");
            expect(reader.readUInt8()).toBe(0xff);
        });

        it("readNullTerminatedString8() stops at the terminator and advances past it.", () => {
            const buf = Buffer.concat([Buffer.from("hi", "utf-8"), Buffer.from([0x00]), Buffer.from([0xff])]);
            const reader = new BufferReader(buf);
            expect(reader.readNullTerminatedString8()).toBe("hi");
            expect(reader.readUInt8()).toBe(0xff);
        });

        it("Reads signed/floating/64-bit primitives.", () => {
            const writer = new BufferWriter();
            writer.writeInt32LE(-5);
            writer.writeFloatLE(2.5);
            writer.writeDoubleLE(3.5);
            writer.writeBigUInt64LE(123n);
            const reader = new BufferReader(writer.toBuffer());
            expect(reader.readInt32LE()).toBe(-5);
            expect(reader.readFloatLE()).toBeCloseTo(2.5, 5);
            expect(reader.readDoubleLE()).toBe(3.5);
            expect(reader.readBigUInt64LE()).toBe(123n);
        });
    });

    describe("BufferWriter", () => {
        it("Concatenates every written chunk in order via toBuffer().", () => {
            const writer = new BufferWriter();
            writer.writeUInt8(1).writeUInt16LE(2).writeUInt32LE(3).writeBytes(Buffer.from([9, 9]));
            const buf = writer.toBuffer();
            expect(buf.length).toBe(1 + 2 + 4 + 2);
            expect(buf[0]).toBe(1);
        });
    });
});
