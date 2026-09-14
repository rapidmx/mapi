///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// AppointmentRecurrence is pure binary-format logic with no DI/DB dependency, tested directly here - same
// precedent as test/mapi/codec/PropertyValue.test.ts.
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { decodeAppointmentRecurrence, encodeAppointmentRecurrence } from "../../src/codec/AppointmentRecurrence.js";
import { RecurrenceFrequency, type RecurrenceRule } from "@rapidmx/restapi";

function roundTrip(rule: RecurrenceRule, startDate: Date, endDate: Date): RecurrenceRule {
    const encoded = encodeAppointmentRecurrence(rule, startDate, endDate);
    return decodeAppointmentRecurrence(new BufferReader(encoded));
}

describe("AppointmentRecurrence Tests", () => {
    describe("Daily", () => {
        it("Round-trips a simple daily recurrence (interval=1).", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            expect(roundTrip(rule, start, end)).toEqual(rule);
        });

        it("Round-trips a multi-day interval (every 3 days).", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 3, exceptions: [] };
            const start = new Date("2026-09-07T08:30:00.000Z");
            const end = new Date("2026-09-07T09:00:00.000Z");
            expect(roundTrip(rule, start, end)).toEqual(rule);
        });
    });

    describe("Weekly", () => {
        it("Round-trips a weekly recurrence with an explicit byDay set.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.WEEKLY, interval: 1, byDay: ["MO", "WE", "FR"], exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z"); // a Monday
            const end = new Date("2026-09-07T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byDay).toEqual(["MO", "WE", "FR"]);
            expect(decoded.freq).toBe(RecurrenceFrequency.WEEKLY);
            expect(decoded.interval).toBe(1);
        });

        it("Defaults byDay to the occurrence's own day-of-week when omitted.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.WEEKLY, interval: 2, exceptions: [] };
            const start = new Date("2026-09-09T10:00:00.000Z"); // a Wednesday
            const end = new Date("2026-09-09T11:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byDay).toEqual(["WE"]);
            expect(decoded.interval).toBe(2);
        });

        it("Handles a week starting on a Sunday (dayOfWeek=0) without underflowing the week-start calculation.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.WEEKLY, interval: 1, byDay: ["SU"], exceptions: [] };
            const start = new Date("2026-09-06T00:00:00.000Z"); // a Sunday
            const end = new Date("2026-09-06T01:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byDay).toEqual(["SU"]);
        });

        it("Ignores an unrecognized byDay code rather than corrupting the mask.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.WEEKLY, interval: 1, byDay: ["MO", "XX"], exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z"); // a Monday
            const end = new Date("2026-09-07T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byDay).toEqual(["MO"]);
        });
    });

    describe("Monthly", () => {
        it("Round-trips a monthly recurrence with an explicit byMonthDay.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.MONTHLY, interval: 1, byMonthDay: [15], exceptions: [] };
            const start = new Date("2026-09-15T14:00:00.000Z");
            const end = new Date("2026-09-15T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byMonthDay).toEqual([15]);
            expect(decoded.freq).toBe(RecurrenceFrequency.MONTHLY);
        });

        it("Defaults byMonthDay to the occurrence's own day-of-month when omitted.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.MONTHLY, interval: 2, exceptions: [] };
            const start = new Date("2026-09-22T14:00:00.000Z");
            const end = new Date("2026-09-22T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.byMonthDay).toEqual([22]);
            expect(decoded.interval).toBe(2);
        });
    });

    describe("Yearly", () => {
        it("Round-trips a plain yearly recurrence (interval=1) using the real RecurFrequency=Yearly wire value.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.YEARLY, interval: 1, exceptions: [] };
            const start = new Date("2026-03-15T14:00:00.000Z");
            const end = new Date("2026-03-15T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.freq).toBe(RecurrenceFrequency.YEARLY);
            expect(decoded.interval).toBe(1);
            expect(decoded.byMonth).toEqual([3]);
            expect(decoded.byMonthDay).toEqual([15]);
        });

        it("Encodes a multi-year yearly recurrence (interval>1) as Monthly/Period=12*N, decoding back as Monthly per the documented asymmetry.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.YEARLY, interval: 2, exceptions: [] };
            const start = new Date("2026-06-01T14:00:00.000Z");
            const end = new Date("2026-06-01T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.freq).toBe(RecurrenceFrequency.MONTHLY);
            expect(decoded.interval).toBe(24);
        });
    });

    describe("End condition", () => {
        it("Round-trips an until (EndAfterDate) end condition.", () => {
            const rule: RecurrenceRule = {
                freq: RecurrenceFrequency.DAILY,
                interval: 1,
                until: new Date("2026-12-01T00:00:00.000Z"),
                exceptions: [],
            };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.until).toEqual(new Date("2026-12-01T00:00:00.000Z"));
            expect(decoded.count).toBeUndefined();
        });

        it("Round-trips a count (EndAfterN) end condition.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, count: 10, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.count).toBe(10);
            expect(decoded.until).toBeUndefined();
        });

        it("Round-trips a never-ending recurrence (no until, no count).", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const decoded = roundTrip(rule, start, end);
            expect(decoded.until).toBeUndefined();
            expect(decoded.count).toBeUndefined();
        });
    });

    describe("Error handling for unsupported wire content", () => {
        it("Throws decoding an unsupported RecurFrequency value.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            encoded.writeUInt16LE(0x9999, 4); // corrupt RecurFrequency (offset 4: after ReaderVersion+WriterVersion)
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(/unsupported RecurFrequency/);
        });

        it("Throws decoding a Daily RecurFrequency paired with a non-Day PatternType.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            encoded.writeUInt16LE(0x0001, 6); // corrupt PatternType (offset 6) to Week
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(/unsupported PatternType/);
        });

        it("Throws decoding a Weekly RecurFrequency paired with a non-Week PatternType.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.WEEKLY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            encoded.writeUInt16LE(0x0002, 6); // corrupt PatternType (offset 6) to Month
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(/unsupported PatternType/);
        });

        it("Throws decoding a Monthly RecurFrequency paired with a non-Month PatternType.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.MONTHLY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            encoded.writeUInt16LE(0x0000, 6); // corrupt PatternType (offset 6) to Day
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(/unsupported PatternType/);
        });

        it("Throws decoding a Yearly RecurFrequency paired with a non-Month PatternType.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.YEARLY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            encoded.writeUInt16LE(0x0000, 6); // corrupt PatternType (offset 6) to Day
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(/unsupported PatternType/);
        });

        it("Decodes deleted occurrences into exceptions instead of throwing, ignoring modified ones and the ExceptionInfo blocks.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            // RecurrencePattern (no PatternTypeSpecific bytes for Day) is 34 bytes up to DeletedInstanceCount.
            const minutes = (iso: string) => Math.round((Date.parse(iso) - Date.UTC(1601, 0, 1)) / 60000);
            const deleted = [minutes("2026-09-09T00:00:00.000Z"), minutes("2026-09-10T00:00:00.000Z")];
            const modified = [minutes("2026-09-10T00:00:00.000Z")];
            const lists = Buffer.alloc(4 + deleted.length * 4 + 4 + modified.length * 4);
            let offset = lists.writeUInt32LE(deleted.length, 0);
            deleted.forEach((value) => (offset = lists.writeUInt32LE(value, offset)));
            offset = lists.writeUInt32LE(modified.length, offset);
            modified.forEach((value) => (offset = lists.writeUInt32LE(value, offset)));
            const withExceptions = Buffer.concat([encoded.subarray(0, 34), lists, encoded.subarray(42)]);
            // ExceptionCount (after StartDate/EndDate and the four outer version/offset fields) says one ExceptionInfo follows.
            withExceptions.writeUInt16LE(1, 34 + lists.length + 8 + 16);

            const decoded = decodeAppointmentRecurrence(new BufferReader(withExceptions));

            expect(decoded.freq).toBe(RecurrenceFrequency.DAILY);
            expect(decoded.exceptions).toEqual([new Date("2026-09-09T14:00:00.000Z")]);
        });

        it("Rejects an instance date count larger than the pattern.", () => {
            const encoded = encodeAppointmentRecurrence({ freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] }, new Date(), new Date());
            encoded.writeUInt32LE(0x40000000, 34);
            expect(() => decodeAppointmentRecurrence(new BufferReader(encoded))).toThrow(RangeError);
        });

        it("Throws encoding an unsupported RecurrenceFrequency value.", () => {
            const rule = { freq: "hourly" as RecurrenceFrequency, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            expect(() => encodeAppointmentRecurrence(rule, start, end)).toThrow(/unsupported RecurrenceFrequency/);
        });
    });

    describe("Reserved blocks", () => {
        it("Skips nonzero-sized ReservedBlock1/ReservedBlock2 content rather than misparsing subsequent fields.", () => {
            const rule: RecurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const start = new Date("2026-09-07T14:00:00.000Z");
            const end = new Date("2026-09-07T15:00:00.000Z");
            const encoded = encodeAppointmentRecurrence(rule, start, end);
            // ReservedBlock1Size(4) sits right after ExceptionCount(2) at offset 66, i.e. offset 68. Rewrite the
            // tail of the buffer to insert 2 padding bytes for a nonzero ReservedBlock1, followed by
            // ReservedBlock2Size(4)=0.
            const head = encoded.subarray(0, 68);
            const reservedBlock1Size = Buffer.alloc(4);
            reservedBlock1Size.writeUInt32LE(2, 0);
            const reservedBlock1 = Buffer.from([0xaa, 0xbb]);
            const reservedBlock2Size = Buffer.alloc(4);
            reservedBlock2Size.writeUInt32LE(0, 0);
            const patched = Buffer.concat([head, reservedBlock1Size, reservedBlock1, reservedBlock2Size]);

            const decoded = decodeAppointmentRecurrence(new BufferReader(patched));
            expect(decoded.freq).toBe(RecurrenceFrequency.DAILY);
        });
    });
});
