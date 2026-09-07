///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// NspiCodec is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// test/mapi/codec/AppointmentRecurrence.test.ts.
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, writePropertyTag, writeTaggedPropertyValue } from "../../src/codec/PropertyValue.js";
import {
    BLANK_STAT,
    extractContentRestrictionSearchTerm,
    readLargePropertyTagArray,
    readStat,
    writeAddressBookPropertyRow,
    writeLargePropertyTagArray,
    writeStat,
    type Stat,
} from "../../src/nspi/NspiCodec.js";

describe("NspiCodec Tests", () => {
    describe("STAT", () => {
        it("Round-trips a STAT structure through write/read.", () => {
            const stat: Stat = {
                sortType: 1,
                containerId: 2,
                currentRec: 3,
                delta: -4,
                numPos: 5,
                totalRecs: 6,
                codePage: 1252,
                templateLocale: 1033,
                sortLocale: 1033,
            };
            const writer = new BufferWriter();
            writeStat(writer, stat);
            const buffer = writer.toBuffer();
            expect(buffer.length).toBe(36);
            expect(readStat(new BufferReader(buffer))).toEqual(stat);
        });

        it("BLANK_STAT is all zeros.", () => {
            const writer = new BufferWriter();
            writeStat(writer, BLANK_STAT);
            expect(writer.toBuffer()).toEqual(Buffer.alloc(36));
        });
    });

    describe("LargePropertyTagArray", () => {
        it("Round-trips an empty array.", () => {
            const writer = new BufferWriter();
            writeLargePropertyTagArray(writer, []);
            const reader = new BufferReader(writer.toBuffer());
            expect(readLargePropertyTagArray(reader)).toEqual([]);
        });

        it("Round-trips a non-empty array of PropertyTags.", () => {
            const tags = [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString },
                { propertyId: 0x3003, propertyType: PropertyType.PtypString },
            ];
            const writer = new BufferWriter();
            writeLargePropertyTagArray(writer, tags);
            const reader = new BufferReader(writer.toBuffer());
            expect(readLargePropertyTagArray(reader)).toEqual(tags);
        });
    });

    describe("writeAddressBookPropertyRow", () => {
        it("Writes Flags=0x00 followed by a bare value (no HasValue byte) for a fixed-size type.", () => {
            const writer = new BufferWriter();
            writeAddressBookPropertyRow(writer, [{ propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 }], [42]);
            const reader = new BufferReader(writer.toBuffer());
            expect(reader.readUInt8()).toBe(0x00); // Flags
            expect(reader.readInt32LE()).toBe(42);
            expect(reader.hasMore()).toBe(false);
        });

        it("Writes Flags=0x00 followed by HasValue=0xFF then the value for a PtypString column.", () => {
            const writer = new BufferWriter();
            writeAddressBookPropertyRow(writer, [{ propertyId: 0x3001, propertyType: PropertyType.PtypString }], ["Jane Doe"]);
            const reader = new BufferReader(writer.toBuffer());
            expect(reader.readUInt8()).toBe(0x00); // Flags
            expect(reader.readUInt8()).toBe(0xff); // HasValue
            expect(reader.readNullTerminatedUtf16LE()).toBe("Jane Doe");
            expect(reader.hasMore()).toBe(false);
        });

        it("Writes multiple columns in order.", () => {
            const writer = new BufferWriter();
            writeAddressBookPropertyRow(
                writer,
                [
                    { propertyId: 0x3001, propertyType: PropertyType.PtypString },
                    { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 },
                ],
                ["Jane", 7],
            );
            const reader = new BufferReader(writer.toBuffer());
            reader.readUInt8(); // Flags
            reader.readUInt8(); // HasValue
            expect(reader.readNullTerminatedUtf16LE()).toBe("Jane");
            expect(reader.readInt32LE()).toBe(7);
        });
    });

    describe("extractContentRestrictionSearchTerm", () => {
        function buildContentRestriction(searchTerm: string): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x03); // RestrictType - ContentRestriction
            writer.writeUInt16LE(0x0001); // FuzzyLevelLow
            writer.writeUInt16LE(0x0001); // FuzzyLevelHigh
            writePropertyTag(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString });
            writeTaggedPropertyValue(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString, value: searchTerm });
            return writer.toBuffer();
        }

        it("Extracts the search term from a well-formed ContentRestriction.", () => {
            const buffer = buildContentRestriction("jane");
            expect(extractContentRestrictionSearchTerm(new BufferReader(buffer))).toBe("jane");
        });

        it("Throws for any restriction type other than ContentRestriction (0x03).", () => {
            const writer = new BufferWriter();
            writer.writeUInt8(0x00); // RestrictType - AndRestriction
            expect(() => extractContentRestrictionSearchTerm(new BufferReader(writer.toBuffer()))).toThrow(/unsupported restriction type/);
        });

        it("Throws when the ContentRestriction's TaggedValue isn't a string.", () => {
            const writer = new BufferWriter();
            writer.writeUInt8(0x03);
            writer.writeUInt16LE(0x0001);
            writer.writeUInt16LE(0x0001);
            writePropertyTag(writer, { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 });
            writeTaggedPropertyValue(writer, { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32, value: 42 });
            expect(() => extractContentRestrictionSearchTerm(new BufferReader(writer.toBuffer()))).toThrow(/TaggedValue was not a string/);
        });
    });
});
