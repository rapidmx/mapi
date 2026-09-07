///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { readTypedString, writeTypedString } from "../../src/codec/TypedString.js";

describe("TypedString Tests", () => {
    it("Encodes undefined as StringType 0x00 (no string) and decodes it back to undefined.", () => {
        const writer = new BufferWriter();
        writeTypedString(writer, undefined);
        const bytes = writer.toBuffer();
        expect(bytes.toString("hex")).toBe("00");
        expect(readTypedString(new BufferReader(bytes))).toBeUndefined();
    });

    it("Encodes an empty string as StringType 0x01 and decodes it back to an empty string.", () => {
        const writer = new BufferWriter();
        writeTypedString(writer, "");
        const bytes = writer.toBuffer();
        expect(bytes.toString("hex")).toBe("01");
        expect(readTypedString(new BufferReader(bytes))).toBe("");
    });

    it("Encodes a non-empty string as StringType 0x04 (Unicode) with UTF-16LE null-terminated content.", () => {
        const writer = new BufferWriter();
        writeTypedString(writer, "Hi");
        const bytes = writer.toBuffer();
        const expectedContent = Buffer.concat([Buffer.from("Hi", "utf16le"), Buffer.from([0x00, 0x00])]);
        expect(bytes[0]).toBe(0x04);
        expect(bytes.subarray(1)).toEqual(expectedContent);
        expect(readTypedString(new BufferReader(bytes))).toBe("Hi");
    });

    it("Decodes a hand-built StringType 0x02 (String8) value.", () => {
        const stringBytes = Buffer.concat([Buffer.from("Hi8", "utf-8"), Buffer.from([0x00])]);
        const bytes = Buffer.concat([Buffer.from([0x02]), stringBytes]);
        expect(readTypedString(new BufferReader(bytes))).toBe("Hi8");
    });

    it("Decodes a hand-built StringType 0x03 (reduced Unicode) value.", () => {
        const bytes = Buffer.from([0x03, 0x41, 0x42, 0x00]); // "AB"
        expect(readTypedString(new BufferReader(bytes))).toBe("AB");
    });
});
