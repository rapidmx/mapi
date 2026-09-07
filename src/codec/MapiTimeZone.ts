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
 * unique zone name) - this codec resolves it to a synthetic-but-valid `Etc/GMT±N` IANA identifier (note the
 * IANA `Etc/GMT` area's own sign convention is POSIX-inverted from common usage: `Etc/GMT+8` means UTC-8, not
 * UTC+8) rounded to the nearest whole hour, since `Etc/GMT` zones only support integer-hour offsets - a further
 * documented approximation for zones with a half-hour/quarter-hour offset (e.g. India, Nepal).
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
    const bias = timezone === "UTC" ? 0 : -utcOffsetMinutes(timezone, at);

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

    if (bias === 0) {
        return "UTC";
    }

    const offsetMinutes = -bias;
    const wholeHours = Math.round(offsetMinutes / 60);
    const etcSign = wholeHours <= 0 ? "+" : "-";
    return `Etc/GMT${etcSign}${Math.abs(wholeHours)}`;
}
