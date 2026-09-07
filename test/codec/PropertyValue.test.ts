///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// PropertyValue is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent
// as test/mapi/codec/MapiGuid.test.ts and test/eas/codec/WbxmlCodec.test.ts.
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import {
    dateToFiletime,
    filetimeToDate,
    PropertyType,
    readPropertyTag,
    readPropertyValue,
    readTaggedPropertyValue,
    readTypedPropertyValue,
    writePropertyTag,
    writePropertyValue,
    writeTaggedPropertyValue,
    writeTypedPropertyValue,
    type PropertyTag,
    type TaggedPropertyValue,
    type TypedPropertyValue,
} from "../../src/codec/PropertyValue.js";

describe("PropertyValue Tests", () => {
    describe("PropertyTag", () => {
        it("Encodes PropertyType in the low 16 bits and PropertyId in the high 16 bits, matching the real PR_SUBJECT constant (0x0037001F).", () => {
            const tag: PropertyTag = { propertyId: 0x0037, propertyType: PropertyType.PtypString };
            const writer = new BufferWriter();
            writePropertyTag(writer, tag);
            const buf = writer.toBuffer();
            expect(buf.readUInt32LE(0)).toBe(0x0037001f);
        });

        it("Round-trips through read/write.", () => {
            const tag: PropertyTag = { propertyId: 0x3007, propertyType: PropertyType.PtypTime };
            const writer = new BufferWriter();
            writePropertyTag(writer, tag);
            const decoded = readPropertyTag(new BufferReader(writer.toBuffer()));
            expect(decoded).toEqual(tag);
        });
    });

    describe("readPropertyValue / writePropertyValue round-trips", () => {
        const roundTrip = (type: PropertyType, value: unknown): unknown => {
            const writer = new BufferWriter();
            writePropertyValue(writer, type, value as any);
            return readPropertyValue(new BufferReader(writer.toBuffer()), type);
        };

        it("PtypInteger16 (signed).", () => {
            expect(roundTrip(PropertyType.PtypInteger16, -1234)).toBe(-1234);
        });

        it("PtypInteger32 (signed).", () => {
            expect(roundTrip(PropertyType.PtypInteger32, -70000)).toBe(-70000);
        });

        it("PtypFloating32.", () => {
            expect(roundTrip(PropertyType.PtypFloating32, 1.5)).toBeCloseTo(1.5, 5);
        });

        it("PtypFloating64.", () => {
            expect(roundTrip(PropertyType.PtypFloating64, 1.23456789)).toBeCloseTo(1.23456789, 10);
        });

        it("PtypBoolean true/false.", () => {
            expect(roundTrip(PropertyType.PtypBoolean, true)).toBe(true);
            expect(roundTrip(PropertyType.PtypBoolean, false)).toBe(false);
        });

        it("PtypInteger64 (bigint).", () => {
            expect(roundTrip(PropertyType.PtypInteger64, 9007199254740993n)).toBe(9007199254740993n);
        });

        it("PtypTime (FILETIME <-> Date), truncated to millisecond precision.", () => {
            const date = new Date("2026-03-15T09:30:00.000Z");
            expect(roundTrip(PropertyType.PtypTime, date)).toEqual(date);
        });

        it("PtypGuid.", () => {
            expect(roundTrip(PropertyType.PtypGuid, "12345678-9abc-def0-1234-56789abcdef0")).toBe(
                "12345678-9abc-def0-1234-56789abcdef0",
            );
        });

        it("PtypString (UTF-16LE, null-terminated), including non-ASCII text.", () => {
            expect(roundTrip(PropertyType.PtypString, "Hello, 世界")).toBe("Hello, 世界");
        });

        it("PtypString8 (UTF-8, null-terminated).", () => {
            expect(roundTrip(PropertyType.PtypString8, "plain ascii")).toBe("plain ascii");
        });

        it("PtypBinary, with a 16-bit length prefix.", () => {
            const data = Buffer.from([1, 2, 3, 4, 5]);
            const result = roundTrip(PropertyType.PtypBinary, data) as Buffer;
            expect(Buffer.compare(result, data)).toBe(0);
        });

        it("PtypMultipleInteger32.", () => {
            expect(roundTrip(PropertyType.PtypMultipleInteger32, [1, -2, 3])).toEqual([1, -2, 3]);
        });

        it("PtypMultipleString.", () => {
            expect(roundTrip(PropertyType.PtypMultipleString, ["a", "bb", "ccc"])).toEqual(["a", "bb", "ccc"]);
        });

        it("PtypMultipleString8.", () => {
            expect(roundTrip(PropertyType.PtypMultipleString8, ["a", "bb"])).toEqual(["a", "bb"]);
        });

        it("PtypMultipleBinary.", () => {
            const values = [Buffer.from([1, 2]), Buffer.from([3, 4, 5])];
            const result = roundTrip(PropertyType.PtypMultipleBinary, values) as Buffer[];
            expect(result.map((b) => b.toString("hex"))).toEqual(values.map((b) => b.toString("hex")));
        });

        it("PtypNull.", () => {
            const writer = new BufferWriter();
            writePropertyValue(writer, PropertyType.PtypNull, 0);
            expect(writer.toBuffer().length).toBe(0);
            expect(readPropertyValue(new BufferReader(Buffer.alloc(0)), PropertyType.PtypNull)).toBe(0);
        });

        it("Throws for an unsupported PropertyType.", () => {
            expect(() => readPropertyValue(new BufferReader(Buffer.alloc(0)), 0x00fd as PropertyType)).toThrow(
                /unsupported PropertyType/i,
            );
            expect(() => writePropertyValue(new BufferWriter(), 0x00fd as PropertyType, 0)).toThrow(
                /unsupported PropertyType/i,
            );
        });
    });

    describe("TypedPropertyValue", () => {
        it("Round-trips the PropertyType tag alongside the value.", () => {
            const typed: TypedPropertyValue = { propertyType: PropertyType.PtypString, value: "Test Subject" };
            const writer = new BufferWriter();
            writeTypedPropertyValue(writer, typed);
            const decoded = readTypedPropertyValue(new BufferReader(writer.toBuffer()));
            expect(decoded).toEqual(typed);
        });
    });

    describe("TaggedPropertyValue", () => {
        it("Round-trips the full PropertyTag (PropertyId and PropertyType) alongside the value.", () => {
            const tagged: TaggedPropertyValue = { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Test Subject" };
            const writer = new BufferWriter();
            writeTaggedPropertyValue(writer, tagged);
            const decoded = readTaggedPropertyValue(new BufferReader(writer.toBuffer()));
            expect(decoded).toEqual(tagged);
        });
    });

    describe("FILETIME conversion", () => {
        it("Converts the Unix epoch to/from the correct FILETIME value.", () => {
            // 1970-01-01T00:00:00Z is exactly 116444736000000000 100ns-intervals after 1601-01-01T00:00:00Z.
            expect(dateToFiletime(new Date(0))).toBe(116444736000000000n);
            expect(filetimeToDate(116444736000000000n)).toEqual(new Date(0));
        });
    });
});
