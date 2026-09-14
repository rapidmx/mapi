///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";

/**
 * Encodes/decodes `CalendarEvent.timezone` (a plain IANA zone identifier) to/from `PidLidTimeZoneStruct`'s
 * `TimeZoneStruct` BLOB (`[MS-OXOCAL]` §2.2.1.39, confirmed field-by-field this session, including a real
 * worked example: `lBias(4,LONG)+lStandardBias(4,LONG)+lDaylightBias(4,LONG)+wStandardYear(2,WORD)+
 * stStandardDate(16,SYSTEMTIME)+wDaylightYear(2,WORD)+stDaylightDate(16,SYSTEMTIME)` = 48 bytes total, all
 * little-endian. `SYSTEMTIME` (`[MS-DTYP]`) is 8 `WORD` fields (`wYear`/`wMonth`/`wDayOfWeek`/`wDay`/`wHour`/
 * `wMinute`/`wSecond`/`wMilliseconds`), 16 bytes.
 *
 * This is deliberately the older, simpler `TimeZoneStruct` - not the newer `TimeZoneDefinition`/
 * `PidLidAppointmentTimeZoneDefinitionStartDisplay`, which additionally encodes named DST transition rules.
 *
 * **Pragmatic scope, matching the same limits `CalendarSyncAdapter` (EAS) already carries for timezones**:
 * - `"UTC"` encodes as an all-zero struct (no bias, no DST) - the spec-legitimate representation of "no offset,
 * no DST".
 * - Any other IANA zone is approximated as a **fixed-offset** zone: `lBias` is derived from that zone's actual
 * UTC offset at the given reference instant (via `Intl`), with `lStandardBias`/`lDaylightBias` always `0` and
 * both `SYSTEMTIME` transition dates all-zero (`wMonth=0`, the spec's own documented way to say "this zone
 * does not observe daylight saving time"). **This is wrong across a real DST boundary** for a zone that
 * actually observes DST - a real client will show the correct offset only for occurrences near the reference
 * instant used at encode time, not a spec violation but a real, documented fidelity gap.
 * - Decoding a struct back into an IANA identifier is fundamentally lossy (a bias alone doesn't identify a
 * unique zone name) - this codec resolves a whole-hour offset to a synthetic-but-valid `Etc/GMT±N` IANA identifier
 * (note the IANA `Etc/GMT` area's own sign convention is POSIX-inverted from common usage: `Etc/GMT+8` means UTC-8,
 * not UTC+8). A half-hour/quarter-hour offset keeps its minutes: it maps to a real DST-less zone with that offset
 * (e.g. `Asia/Kolkata`, `Asia/Kathmandu`) or otherwise to a fixed-offset identifier like `"-03:30"`. An offset
 * outside UTC-12..UTC+14 decodes as `"UTC"`.
 * - Encoding a zone `Intl` doesn't recognize falls back to UTC instead of throwing.
 *
 * @author Jean-Philippe Steinmetz
 */

const SYSTEMTIME_FIELD_COUNT = 8;

function writeZeroTransitionBlock(writer: BufferWriter): void {
    writer.writeUInt16LE(0); // wStandardYear / wDaylightYear
    for (let i = 0; i < SYSTEMTIME_FIELD_COUNT; i++) {
        writer.writeUInt16LE(0);
    }
}

function skipTransitionBlock(reader: BufferReader): void {
    reader.readUInt16LE(); // wStandardYear / wDaylightYear
    for (let i = 0; i < SYSTEMTIME_FIELD_COUNT; i++) {
        reader.readUInt16LE();
    }
}

/** The zone's UTC offset in minutes at instant `at` (standard convention: local minus UTC, e.g. `-480` for
 * `America/Los_Angeles` in winter), read via `Intl`'s `shortOffset` formatting rather than hand-rolling a
 * zone database. Throws the same way `Intl.DateTimeFormat` itself does for an unrecognized zone identifier. */
function utcOffsetMinutes(timezone: string, at: Date): number {
    const formatter = new Intl.DateTimeFormat("en-US", { timeZone: timezone, timeZoneName: "shortOffset" });
    // `formatToParts` always includes a "timeZoneName" part when `timeZoneName` is requested in the
    // formatter's own options, so the non-null assertion below reflects a real `Intl` guarantee, not an
    // unverified assumption.
    const tzName = formatter.formatToParts(at).find((part) => part.type === "timeZoneName")!.value;
    const match = /^GMT([+-])(\d+)(?::(\d+))?$/.exec(tzName);
    if (!match) {
        return 0; // Bare "GMT", i.e. UTC itself.
    }
    const sign = match[1] === "-" ? -1 : 1;
    const hours = parseInt(match[2], 10);
    const minutes = match[3] ? parseInt(match[3], 10) : 0;
    return sign * (hours * 60 + minutes);
}

/** Encodes `timezone` (an IANA identifier, or `"UTC"`) into a 48-byte `TimeZoneStruct` BLOB, using `at` as the
 * reference instant for resolving a non-UTC zone's current fixed offset. */
export function encodeTimeZoneStruct(timezone: string, at: Date): Buffer {
    const writer = new BufferWriter();
    let bias = 0;
    try {
        bias = timezone === "UTC" ? 0 : -utcOffsetMinutes(timezone, at);
    } catch {
        // An identifier Intl doesn't recognize (e.g. stored by another client) is sent as UTC rather than failing
        // the whole property read.
        bias = 0;
    }

    writer.writeInt32LE(bias); // lBias
    writer.writeInt32LE(0); // lStandardBias - no DST modeled
    writer.writeInt32LE(0); // lDaylightBias - no DST modeled
    writeZeroTransitionBlock(writer); // wStandardYear + stStandardDate
    writeZeroTransitionBlock(writer); // wDaylightYear + stDaylightDate

    return writer.toBuffer();
}

/** Decodes a `TimeZoneStruct` BLOB (read from `reader`'s current position) into an approximate IANA zone
 * identifier - see this file's own doc comment for the lossy `Etc/GMT±N` fallback's exact rounding rules. */
export function decodeTimeZoneStruct(reader: BufferReader): string {
    const bias = reader.readInt32LE();
    reader.readInt32LE(); // lStandardBias - not reconstructable into a real DST rule, ignored
    reader.readInt32LE(); // lDaylightBias - ditto
    skipTransitionBlock(reader); // wStandardYear + stStandardDate
    skipTransitionBlock(reader); // wDaylightYear + stDaylightDate

    const offsetMinutes = -bias;
    // No real zone is further than UTC-12/UTC+14; anything else is garbage and treated as UTC.
    if (offsetMinutes === 0 || offsetMinutes < -12 * 60 || offsetMinutes > 14 * 60) {
        return "UTC";
    }
    if (offsetMinutes % 60 === 0) {
        const hours = offsetMinutes / 60;
        return `Etc/GMT${hours < 0 ? "+" : "-"}${Math.abs(hours)}`;
    }
    return FRACTIONAL_OFFSET_ZONES[offsetMinutes] ?? formatFixedOffset(offsetMinutes);
}

/** Real zones with a non-whole-hour offset and no daylight saving, so the zone matches a DST-less
 * `TimeZoneStruct` all year. */
const FRACTIONAL_OFFSET_ZONES: Record<number, string> = {
    [-570]: "Pacific/Marquesas",
    210: "Asia/Tehran",
    270: "Asia/Kabul",
    330: "Asia/Kolkata",
    345: "Asia/Kathmandu",
    390: "Asia/Yangon",
    525: "Australia/Eucla",
    570: "Australia/Darwin",
};

/** A fixed-offset identifier such as `"-03:30"`, which `Intl.DateTimeFormat` accepts as a `timeZone`. Used for a
 * minute-precision offset with no matching DST-less zone, instead of rounding it to a whole hour. */
function formatFixedOffset(offsetMinutes: number): string {
    const magnitude = Math.abs(offsetMinutes);
    const hours = String(Math.floor(magnitude / 60)).padStart(2, "0");
    const minutes = String(magnitude % 60).padStart(2, "0");
    return `${offsetMinutes < 0 ? "-" : "+"}${hours}:${minutes}`;
}
