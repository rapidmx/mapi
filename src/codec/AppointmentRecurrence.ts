///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";
import { RecurrenceFrequency, type RecurrenceRule } from "@rapidmx/restapi";

/**
 * Encodes/decodes `CalendarEvent.recurrenceRule` to/from the binary blob carried by `PidLidAppointmentRecur`
 * (`[MS-OXOCAL]` §2.2.1.44: the outer `AppointmentRecurrencePattern` structure, §2.2.1.44.5, wrapping an inner
 * `RecurrencePattern`, §2.2.1.44.1). Every field layout/size below was confirmed against those spec pages
 * directly, not assumed.
 *
 * **Pragmatic scope, matching the same limits `CalendarSyncAdapter` (EAS) already carries**:
 * - Only `RecurFrequency` Daily/Weekly/Monthly/Yearly and `PatternType` Day/Week/Month are produced or
 * accepted - `MonthNth`/`MonthEnd`/Hijri patterns have no representation in `RecurrenceRule` (no ordinal-
 * weekday field) and decoding one throws a clear, documented error rather than silently guessing.
 * - `DeletedInstanceCount`/`ModifiedInstanceCount` (recurrence exceptions) and the outer structure's
 * `ExceptionCount` are always written as `0`. Decoding reads the deleted occurrences into `RecurrenceRule.exceptions`
 * and ignores the details of modified ones (`ExceptionInfo`), which `RecurrenceRule` can't hold, instead of failing
 * the whole save.
 * - A `YEARLY` rule with `interval > 1` ("every N years", N>1) is encoded as `RecurFrequency=Monthly` with
 * `Period=12*N`, per the spec's own stated rule that "a yearly recurrence pattern is just a monthly pattern
 * that occurs every 12 months" (a plain `RecurFrequency=Yearly` value's `Period` MUST be exactly 12).
 * **This is a one-way, documented round-trip asymmetry**: decoding that same blob back yields
 * `{freq: MONTHLY, interval: 12*N}`, not the original `{freq: YEARLY, interval: N}` - the two are
 * indistinguishable on the wire once encoded, the same ambiguity a real Exchange server itself has no way
 * around either.
 * - `StartTimeOffset`/`EndTimeOffset` (the outer structure) are derived from `startDate`/`endDate`'s own
 * time-of-day only (`hours*60+minutes`), not carried through decode - the real start/end instant of any
 * given occurrence comes from `PidLidAppointmentStartWhole`/`EndWhole`, not from this blob.
 *
 * @author Jean-Philippe Steinmetz
 */

/** `RecurFrequency` (`[MS-OXOCAL]` §2.2.1.44.1). */
enum RecurFrequencyWire {
    Daily = 0x200a,
    Weekly = 0x200b,
    Monthly = 0x200c,
    Yearly = 0x200d,
}

/** `PatternType` (`[MS-OXOCAL]` §2.2.1.44.1) - only the 3 values this codec produces/accepts. */
enum PatternTypeWire {
    Day = 0x0000,
    Week = 0x0001,
    Month = 0x0002,
}

/** `EndType` (`[MS-OXOCAL]` §2.2.1.44.1). */
enum EndTypeWire {
    AfterDate = 0x00002021,
    AfterN = 0x00002022,
    Never = 0x00002023,
}

/** Sentinel `EndDate` value (`[MS-OXOCAL]` §2.2.1.44.1) meaning "never-ending". */
const NEVER_END_DATE = 0x5ae980df;

const MINUTES_PER_DAY = 1440;
const MINUTES_PER_WEEK = 10080;

/** `RecurrencePattern.CalendarType` `Default` value (Gregorian) - the only calendar type this codec supports. */
const CALENDAR_TYPE_DEFAULT = 0x0000;

/** `RecurrencePattern.FirstDOW` `Sunday` value - this codec always encodes with a Sunday week start, since
 * `RecurrenceRule` has no field of its own to source a different one from. */
const FIRST_DOW_SUNDAY = 0x00000000;

/** `RecurrencePattern`'s `DayOfWeekMask` bits (`[MS-OXOCAL]` §2.2.1.44.1) - confirmed identical to
 * `CalendarSyncAdapter`'s own `DAY_OF_WEEK_BITS` table for MS-ASCAL's `DayOfWeek` field, reused here under the
 * same RFC 5545 two-letter day codes `RecurrenceRule.byDay` uses. */
const DAY_OF_WEEK_BITS: Record<string, number> = { SU: 1, MO: 2, TU: 4, WE: 8, TH: 16, FR: 32, SA: 64 };
const DAY_OF_WEEK_CODE_BY_BIT: [string, number][] = [
    ["SU", 1],
    ["MO", 2],
    ["TU", 4],
    ["WE", 8],
    ["TH", 16],
    ["FR", 32],
    ["SA", 64],
];

/** Minutes-since-1601-01-01T00:00:00Z epoch, as a JS `Date`-compatible millisecond timestamp - every date field
 * in `RecurrencePattern` uses this epoch/unit, distinct from `PropertyValue.ts`'s `FILETIME` epoch (which is
 * also 1601-based, but in 100-ns units, not minutes). */
const EPOCH_1601_MS = Date.UTC(1601, 0, 1);

function dateToMinutesSince1601(date: Date): number {
    return Math.round((date.getTime() - EPOCH_1601_MS) / 60000);
}

function minutesSince1601ToDate(minutes: number): Date {
    return new Date(EPOCH_1601_MS + minutes * 60000);
}

function utcMidnight(date: Date): Date {
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

/** Calendar months elapsed between 1601-01-01 and the first day of the month containing `date`. */
function monthsElapsedSince1601(date: Date): number {
    return (date.getUTCFullYear() - 1601) * 12 + date.getUTCMonth();
}

function dayOfWeekMaskFromByDay(byDay: string[] | undefined, fallbackDate: Date): number {
    if (byDay && byDay.length > 0) {
        return byDay.reduce((mask, day) => mask | (DAY_OF_WEEK_BITS[day] ?? 0), 0);
    }
    return DAY_OF_WEEK_CODE_BY_BIT[fallbackDate.getUTCDay()][1];
}

function byDayFromMask(mask: number): string[] {
    return DAY_OF_WEEK_CODE_BY_BIT.filter(([, bit]) => (mask & bit) !== 0).map(([code]) => code);
}

/**
 * Encodes `rule` (plus the occurrence's own `startDate`/`endDate`, needed for phase/time-of-day derivation)
 * into an `AppointmentRecurrencePattern` blob, ready to be written as the raw content of a `PtypBinary`
 * `PidLidAppointmentRecur` property value.
 */
export function encodeAppointmentRecurrence(rule: RecurrenceRule, startDate: Date, endDate: Date): Buffer {
    let recurFrequency: RecurFrequencyWire;
    let patternType: PatternTypeWire;
    let period: number;
    let patternSpecific: Buffer;
    let firstDateTime: number;

    switch (rule.freq) {
        case RecurrenceFrequency.DAILY: {
            recurFrequency = RecurFrequencyWire.Daily;
            patternType = PatternTypeWire.Day;
            period = rule.interval * MINUTES_PER_DAY;
            patternSpecific = Buffer.alloc(0);
            firstDateTime = dateToMinutesSince1601(utcMidnight(startDate)) % period;
            break;
        }
        case RecurrenceFrequency.WEEKLY: {
            recurFrequency = RecurFrequencyWire.Weekly;
            patternType = PatternTypeWire.Week;
            period = rule.interval;
            const mask = dayOfWeekMaskFromByDay(rule.byDay, startDate);
            patternSpecific = new BufferWriter().writeUInt32LE(mask).toBuffer();
            const midnight = utcMidnight(startDate);
            const weekStart = new Date(midnight.getTime() - midnight.getUTCDay() * MINUTES_PER_DAY * 60000);
            firstDateTime = dateToMinutesSince1601(weekStart) % (period * MINUTES_PER_WEEK);
            break;
        }
        case RecurrenceFrequency.MONTHLY: {
            recurFrequency = RecurFrequencyWire.Monthly;
            patternType = PatternTypeWire.Month;
            period = rule.interval;
            const day = rule.byMonthDay?.[0] ?? startDate.getUTCDate();
            patternSpecific = new BufferWriter().writeUInt32LE(day).toBuffer();
            const monthsElapsed = monthsElapsedSince1601(startDate);
            const refMonthIndex = monthsElapsed % period;
            firstDateTime = dateToMinutesSince1601(new Date(Date.UTC(1601, refMonthIndex, 1)));
            break;
        }
        case RecurrenceFrequency.YEARLY: {
            patternType = PatternTypeWire.Month;
            if (rule.interval <= 1) {
                recurFrequency = RecurFrequencyWire.Yearly;
                period = 12;
            } else {
                recurFrequency = RecurFrequencyWire.Monthly;
                period = 12 * rule.interval;
            }
            const day = rule.byMonthDay?.[0] ?? startDate.getUTCDate();
            patternSpecific = new BufferWriter().writeUInt32LE(day).toBuffer();
            const monthsElapsed = monthsElapsedSince1601(startDate);
            const refMonthIndex = monthsElapsed % period;
            firstDateTime = dateToMinutesSince1601(new Date(Date.UTC(1601, refMonthIndex, 1)));
            break;
        }
        default:
            throw new Error(`AppointmentRecurrence: unsupported RecurrenceFrequency ${rule.freq as string}`);
    }

    let endType: EndTypeWire;
    let occurrenceCount = 0;
    let endDateMinutes: number;
    if (rule.until) {
        endType = EndTypeWire.AfterDate;
        endDateMinutes = dateToMinutesSince1601(utcMidnight(rule.until));
    } else if (rule.count != null) {
        endType = EndTypeWire.AfterN;
        occurrenceCount = rule.count;
        endDateMinutes = NEVER_END_DATE;
    } else {
        endType = EndTypeWire.Never;
        endDateMinutes = NEVER_END_DATE;
    }

    const writer = new BufferWriter();
    // RecurrencePattern (MS-OXOCAL §2.2.1.44.1)
    writer.writeUInt16LE(0x3004); // ReaderVersion
    writer.writeUInt16LE(0x3004); // WriterVersion
    writer.writeUInt16LE(recurFrequency);
    writer.writeUInt16LE(patternType);
    writer.writeUInt16LE(CALENDAR_TYPE_DEFAULT);
    writer.writeUInt32LE(firstDateTime);
    writer.writeUInt32LE(period);
    writer.writeUInt32LE(0); // SlidingFlag
    writer.writeBytes(patternSpecific);
    writer.writeUInt32LE(endType);
    writer.writeUInt32LE(occurrenceCount);
    writer.writeUInt32LE(FIRST_DOW_SUNDAY);
    writer.writeUInt32LE(0); // DeletedInstanceCount
    writer.writeUInt32LE(0); // ModifiedInstanceCount
    writer.writeUInt32LE(dateToMinutesSince1601(utcMidnight(startDate))); // StartDate
    writer.writeUInt32LE(endDateMinutes); // EndDate

    // AppointmentRecurrencePattern's own wrapper fields (MS-OXOCAL §2.2.1.44.5)
    writer.writeUInt32LE(0x00003006); // ReaderVersion2
    writer.writeUInt32LE(0x00003009); // WriterVersion2
    writer.writeUInt32LE(startDate.getUTCHours() * 60 + startDate.getUTCMinutes()); // StartTimeOffset
    writer.writeUInt32LE(endDate.getUTCHours() * 60 + endDate.getUTCMinutes()); // EndTimeOffset
    writer.writeUInt16LE(0); // ExceptionCount
    writer.writeUInt32LE(0); // ReservedBlock1Size
    writer.writeUInt32LE(0); // ReservedBlock2Size

    return writer.toBuffer();
}

/** Reads a `DeletedInstanceCount`/`ModifiedInstanceCount`-style count followed by that many 4-byte dates. */
function readDateList(reader: BufferReader): number[] {
    const count = reader.readUInt32LE();
    if (count * 4 > reader.remaining) {
        throw new RangeError(`AppointmentRecurrence: an instance date count of ${count} is larger than the pattern.`);
    }
    return Array.from({ length: count }, () => reader.readUInt32LE());
}

/**
 * Decodes an `AppointmentRecurrencePattern` blob (read from `reader`'s current position) back into a
 * `RecurrenceRule`. Throws a clear error for any pattern this pragmatic subset can't represent (see this
 * file's own doc comment) rather than silently dropping data.
 */
export function decodeAppointmentRecurrence(reader: BufferReader): RecurrenceRule {
    reader.readUInt16LE(); // ReaderVersion
    reader.readUInt16LE(); // WriterVersion
    const recurFrequency: RecurFrequencyWire = reader.readUInt16LE();
    const patternType: PatternTypeWire = reader.readUInt16LE();
    reader.readUInt16LE(); // CalendarType
    const firstDateTime = reader.readUInt32LE();
    const period = reader.readUInt32LE();
    reader.readUInt32LE(); // SlidingFlag

    let freq: RecurrenceFrequency;
    let interval: number;
    let byDay: string[] | undefined;
    let byMonthDay: number[] | undefined;
    let byMonth: number[] | undefined;

    switch (recurFrequency) {
        case RecurFrequencyWire.Daily:
            if (patternType !== PatternTypeWire.Day) {
                throw new Error(`AppointmentRecurrence: unsupported PatternType 0x${patternType.toString(16)} for Daily`);
            }
            freq = RecurrenceFrequency.DAILY;
            interval = period / MINUTES_PER_DAY;
            break;
        case RecurFrequencyWire.Weekly: {
            if (patternType !== PatternTypeWire.Week) {
                throw new Error(`AppointmentRecurrence: unsupported PatternType 0x${patternType.toString(16)} for Weekly`);
            }
            const mask = reader.readUInt32LE();
            freq = RecurrenceFrequency.WEEKLY;
            interval = period;
            byDay = byDayFromMask(mask);
            break;
        }
        case RecurFrequencyWire.Monthly: {
            if (patternType !== PatternTypeWire.Month) {
                throw new Error(`AppointmentRecurrence: unsupported PatternType 0x${patternType.toString(16)} for Monthly`);
            }
            const day = reader.readUInt32LE();
            freq = RecurrenceFrequency.MONTHLY;
            interval = period;
            byMonthDay = [day];
            break;
        }
        case RecurFrequencyWire.Yearly: {
            if (patternType !== PatternTypeWire.Month) {
                throw new Error(`AppointmentRecurrence: unsupported PatternType 0x${patternType.toString(16)} for Yearly`);
            }
            const day = reader.readUInt32LE();
            freq = RecurrenceFrequency.YEARLY;
            interval = 1;
            byMonthDay = [day];
            byMonth = [minutesSince1601ToDate(firstDateTime).getUTCMonth() + 1];
            break;
        }
        default:
            throw new Error(`AppointmentRecurrence: unsupported RecurFrequency 0x${(recurFrequency as number).toString(16)}`);
    }

    const endType: EndTypeWire = reader.readUInt32LE();
    const occurrenceCount = reader.readUInt32LE();
    reader.readUInt32LE(); // FirstDOW
    // DeletedInstanceDates holds the original dates of every deleted *or* modified occurrence; ModifiedInstanceDates
    // the (new) dates of the modified ones. An occurrence deleted outright is in the first list only.
    const deletedDates = readDateList(reader);
    const modifiedDates = new Set(readDateList(reader));
    reader.readUInt32LE(); // StartDate
    const endDateMinutes = reader.readUInt32LE();

    let until: Date | undefined;
    let count: number | undefined;
    if (endType === EndTypeWire.AfterDate) {
        until = minutesSince1601ToDate(endDateMinutes);
    } else if (endType === EndTypeWire.AfterN) {
        count = occurrenceCount;
    }

    reader.readUInt32LE(); // ReaderVersion2
    reader.readUInt32LE(); // WriterVersion2
    const startTimeOffset = reader.readUInt32LE();
    // Everything after this (ExceptionCount, the variable-length ExceptionInfo/ExtendedException blocks for modified
    // occurrences, and the reserved blocks) describes per-occurrence changes RecurrenceRule has no field for, so it is
    // not read. A modified occurrence that stayed on its day remains an ordinary occurrence; one moved to another day
    // only drops out of its original day, since its new time has nowhere to go.

    return {
        freq,
        interval,
        ...(byDay ? { byDay } : {}),
        ...(byMonthDay ? { byMonthDay } : {}),
        ...(byMonth ? { byMonth } : {}),
        ...(until ? { until } : {}),
        ...(count != null ? { count } : {}),
        // A deleted occurrence's date is midnight (minutes since 1601); its start is that plus the series' start time.
        exceptions: deletedDates.filter((minutes) => !modifiedDates.has(minutes)).map((minutes) => minutesSince1601ToDate(minutes + startTimeOffset)),
    };
}
