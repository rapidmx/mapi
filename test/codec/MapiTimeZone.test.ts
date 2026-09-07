///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// MapiTimeZone is pure binary-format logic with no DI/DB dependency, tested directly here - same precedent as
// test/mapi/codec/AppointmentRecurrence.test.ts.
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { decodeTimeZoneStruct, encodeTimeZoneStruct } from "../../src/codec/MapiTimeZone.js";

describe("MapiTimeZone Tests", () => {
    describe("UTC", () => {
        it("Encodes UTC as a 48-byte, all-zero-except-nothing struct.", () => {
            const encoded = encodeTimeZoneStruct("UTC", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.length).toBe(48);
            expect(encoded.every((byte) => byte === 0)).toBe(true);
        });

        it("Round-trips UTC back to the string 'UTC'.", () => {
            const encoded = encodeTimeZoneStruct("UTC", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("UTC");
        });
    });

    describe("Fixed-offset IANA zones", () => {
        it("Encodes America/Los_Angeles in January (PST, UTC-8) with lBias=480.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(480);
            expect(encoded.length).toBe(48);
        });

        it("Encodes America/Los_Angeles in July (PDT, UTC-7) with lBias=420, reflecting the reference instant's own offset.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-07-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(420);
        });

        it("Encodes Asia/Shanghai (UTC+8) with a negative lBias.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Shanghai", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(-480);
        });

        it("Writes zero lStandardBias/lDaylightBias and all-zero transition blocks (no DST modeled).", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(4)).toBe(0); // lStandardBias
            expect(encoded.readInt32LE(8)).toBe(0); // lDaylightBias
            expect(encoded.subarray(12).every((byte) => byte === 0)).toBe(true); // both transition blocks
        });

        it("Decodes an America/Los_Angeles-derived struct back to the Etc/GMT+8 approximation.", () => {
            const encoded = encodeTimeZoneStruct("America/Los_Angeles", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Etc/GMT+8");
        });

        it("Decodes an Asia/Shanghai-derived struct back to the Etc/GMT-8 approximation.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Shanghai", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Etc/GMT-8");
        });

        it("Rounds a half-hour-offset zone (Asia/Kolkata, UTC+5:30) to the nearest whole hour on decode.", () => {
            const encoded = encodeTimeZoneStruct("Asia/Kolkata", new Date("2026-01-15T00:00:00.000Z"));
            expect(decodeTimeZoneStruct(new BufferReader(encoded))).toBe("Etc/GMT-6");
        });

        it("Encodes a zero-offset zone that isn't literally named 'UTC' (Africa/Abidjan) with lBias=0, exercising the bare-'GMT' offset-string branch.", () => {
            const encoded = encodeTimeZoneStruct("Africa/Abidjan", new Date("2026-01-15T00:00:00.000Z"));
            expect(encoded.readInt32LE(0)).toBe(0);
        });
    });

    describe("Error handling", () => {
        it("Throws encoding an unrecognized IANA zone identifier.", () => {
            expect(() => encodeTimeZoneStruct("Not/AZone", new Date("2026-01-15T00:00:00.000Z"))).toThrow();
        });
    });
});
