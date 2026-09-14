///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import MailComposer from "nodemailer/lib/mail-composer/index.js";
import { AttendeeResponseStatus, buildEventIcs, type Attendee, type CalendarEvent } from "@rapidmx/restapi";
import { BufferReader } from "../codec/BufferCursor.js";
import { isPlainEmailAddress } from "./AddressList.js";
import { decodeGlobalObjectId, globalObjectIdInstanceDate } from "../codec/GlobalObjectId.js";
import type { MapiSessionContext } from "../MapiSessionManager.js";
import { LID_GLOBAL_OBJECT_ID, PSETID_MEETING } from "./CalendarNamedProperties.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import { asEntity, boundIndexedValue, literalQueryValue } from "./RestapiRules.js";
import type { RopContext } from "./RopHandler.js";
import { sendOrThrow } from "./TransportSend.js";

/** `PidTagMessageClass` suffix -> the caller's own new `AttendeeResponseStatus`, per `[MS-OXOCAL]`'s meeting
 * response object naming convention. An unrecognized suffix (a message class this pragmatic subset doesn't
 * know how to interpret as a response) is answered with `MAPI_E_INVALID_PARAMETER`. */
const RESPONSE_STATUS_BY_MESSAGE_CLASS: Record<string, AttendeeResponseStatus> = {
    "IPM.Schedule.Meeting.Resp.Pos": AttendeeResponseStatus.ACCEPTED,
    "IPM.Schedule.Meeting.Resp.Tent": AttendeeResponseStatus.TENTATIVE,
    "IPM.Schedule.Meeting.Resp.Neg": AttendeeResponseStatus.DECLINED,
};

/** Scans `properties` (a draft handle's accumulated `RopSetProperties` values) for the one already-assigned
 * named property matching `(guid, lid)`, returning its raw stored string value - the single-property analog of
 * `RopSaveChangesMessageHandler.ts`'s own full-table decode loop, needed here only for `PidLidGlobalObjectId`. */
function findNamedPropertyValue(session: MapiSessionContext, properties: Record<string, string>, guid: string, lid: number): string | undefined {
    for (const [key, value] of Object.entries(properties)) {
        const propertyId = Number(key);
        if (propertyId < 0x8000) {
            continue;
        }
        const named = resolveNamedProperty(session, propertyId);
        if (named && named.kind === "lid" && named.lid === lid && named.guid.toLowerCase() === guid) {
            return value;
        }
    }
    return undefined;
}

/** `MAPI_E_INVALID_PARAMETER`: not a response this server understands (unknown message class, no or malformed
 * `PidLidGlobalObjectId`). */
const ERROR_INVALID_PARAMETER = 0x80070057;
/** `MAPI_E_NOT_FOUND`: no meeting in the caller's mailbox matches, or the caller isn't one of its attendees. */
const ERROR_NOT_FOUND = 0x8004010f;
/** `MAPI_E_CALL_FAILED`: the response was recorded but the iTIP `REPLY` couldn't be sent to the organizer. */
const ERROR_CALL_FAILED = 0x80004005;

/**
 * Handles a submitted `"IPM.Schedule.Meeting.Resp.{Pos,Neg,Tent}"` message - an attendee's own response to a
 * meeting invite. Delegated to from `RopSubmitMessageHandler` for that message-class prefix, instead of the ordinary
 * mail or Appointment paths. Returns the ROP's `ReturnValue`.
 *
 * Mirrors restapi's `BaseCalendarEventRoute.respond()`. It correlates the response via `PidLidGlobalObjectId`
 * (echoed by a real client from the invite it's responding to - see `GlobalObjectId.ts`) to **the caller's own
 * copy** of the meeting: the query is scoped to the caller's `mailboxUid`, which is the access check. The
 * organizer's copy (in another mailbox) is never touched here; the organizer's server applies the iTIP `REPLY` sent
 * below, exactly as for a REST response. A native Outlook `GlobalObjectId` decodes to the uppercase hex form of its
 * iCalendar UID, which is matched in either case.
 *
 * **One occurrence of a recurring meeting** (the `GlobalObjectId`'s instance date is set): the caller's exception
 * copy for that date is used when there is one, comparing dates in the event's own time zone. Without one, the
 * series is left alone - responding to one occurrence must not change or delete every other occurrence. Declining
 * adds the occurrence to the caller's series `recurrenceRule.exceptions` (so it leaves their calendar); accepting or
 * tentatively accepting changes nothing locally. Either way the `REPLY` names just that occurrence (`RECURRENCE-ID`).
 *
 * **The whole meeting** (no instance date, or a copy that doesn't recur): the caller's `Attendee.responseStatus` is
 * recorded; declining soft-deletes the caller's copy instead. Declining an occurrence through its exception copy also
 * adds that occurrence to the series' exceptions, so the series doesn't show it again. Every update is versioned
 * (`asEntity`), so a concurrent change fails the response instead of being overwritten.
 *
 * The `REPLY` (restapi's own `buildEventIcs`) goes out through `sendOrThrow`, so a transport that rejects it is
 * reported as `MAPI_E_CALL_FAILED` instead of success; the response itself stays recorded. An organizer or attendee
 * address that isn't a plain SMTP address is not mailed at all. A response that can't be matched to a meeting the
 * caller attends is an error (`MAPI_E_NOT_FOUND`/`MAPI_E_INVALID_PARAMETER`), not a silent success.
 */
export async function submitMeetingResponse(messageClass: string, draftProperties: Record<string, string>, context: RopContext): Promise<number> {
    const responseStatus = RESPONSE_STATUS_BY_MESSAGE_CLASS[messageClass];
    const globalObjectIdBase64 = responseStatus
        ? findNamedPropertyValue(context.session, draftProperties, PSETID_MEETING, LID_GLOBAL_OBJECT_ID)
        : undefined;
    if (!globalObjectIdBase64) {
        return ERROR_INVALID_PARAMETER;
    }
    const globalObjectId = Buffer.from(globalObjectIdBase64, "base64");
    let icalUid: string;
    try {
        icalUid = decodeGlobalObjectId(new BufferReader(globalObjectId));
    } catch {
        return ERROR_INVALID_PARAMETER;
    }

    const copies = await findCallerCopies(context, icalUid);
    const picked = pickOccurrence(copies, globalObjectIdInstanceDate(globalObjectId));
    const mailbox = picked ? await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true }) : undefined;
    if (!picked || !mailbox) {
        return ERROR_NOT_FOUND;
    }
    const { event, occurrence } = picked;
    const callerAddresses = new Set([mailbox.primarySmtpAddress.toLowerCase(), ...mailbox.aliasAddresses.map((a: string) => a.toLowerCase())]);
    const respondingAttendee = event.attendees.find((attendee) => callerAddresses.has(attendee.address.toLowerCase()));
    if (!respondingAttendee) {
        return ERROR_NOT_FOUND;
    }

    const updatedAttendee: Attendee = { ...respondingAttendee, responseStatus };
    let replyEvent: CalendarEventRow = event;
    if (occurrence) {
        if (responseStatus === AttendeeResponseStatus.DECLINED) {
            await addSeriesException(context, event, occurrence);
        }
        const duration = new Date(event.endDate).getTime() - new Date(event.startDate).getTime();
        replyEvent = { ...event, recurrenceRule: undefined, recurrenceId: occurrence, startDate: occurrence, endDate: new Date(occurrence.getTime() + duration) };
    } else if (responseStatus === AttendeeResponseStatus.DECLINED) {
        await context.calendarEventRepo.delete(event.uid, { ignoreACL: true });
        // Declining an occurrence that has its own exception copy: deleting the copy alone would bring the series'
        // original occurrence back, so it is excluded from the series too (as restapi's ScanQueueJob does for an
        // occurrence CANCEL).
        const series = copies.find((copy) => copy.recurrenceId == null);
        if (event.recurrenceId != null && series?.recurrenceRule) {
            await addSeriesException(context, series, new Date(event.recurrenceId));
        }
    } else {
        const attendees: Attendee[] = event.attendees.map((attendee) => (attendee === respondingAttendee ? updatedAttendee : attendee));
        await context.calendarEventRepo.update({ uid: event.uid, version: event.version, attendees }, asEntity(context.calendarEventRepo, event), {
            ignoreACL: true,
        });
    }

    try {
        await sendReply(context, replyEvent, updatedAttendee, mailbox.displayName);
    } catch {
        return ERROR_CALL_FAILED;
    }
    return 0;
}

/** The most copies of one meeting (series plus exceptions) considered in the caller's mailbox. */
const MAX_EVENT_COPIES = 100;

type CalendarEventRow = CalendarEvent & { uid: string; version: number };

const REPLY_SUBJECT_PREFIX: Record<string, string> = {
    [AttendeeResponseStatus.ACCEPTED]: "Accepted",
    [AttendeeResponseStatus.TENTATIVE]: "Tentative",
    [AttendeeResponseStatus.DECLINED]: "Declined",
};

/** Adds `occurrence` to `series`' recurrence exceptions (a versioned update), unless it is already there. */
async function addSeriesException(context: RopContext, series: CalendarEventRow, occurrence: Date): Promise<void> {
    const rule = series.recurrenceRule!;
    const exceptions = (rule.exceptions ?? []).map((date) => new Date(date));
    if (exceptions.some((date) => date.getTime() === occurrence.getTime())) {
        return;
    }
    const recurrenceRule = { ...rule, exceptions: [...exceptions, occurrence] };
    await context.calendarEventRepo.update({ uid: series.uid, version: series.version, recurrenceRule }, asEntity(context.calendarEventRepo, series), {
        ignoreACL: true,
    });
}

/**
 * The caller's copies of the meeting `icalUid` names. An `OutlookID`'s hex UID may have been stored in either case, so a
 * miss is retried in lower case.
 *
 * The UID comes from the client (and originally from whoever sent the invite), so it is looked up the way restapi
 * stores it (`boundIndexedValue`, hashed past 255 characters), matched literally (a UID like `ne(x)` is not an
 * operator), and every returned row's `icalUid` is compared again.
 */
async function findCallerCopies(context: RopContext, icalUid: string): Promise<CalendarEventRow[]> {
    const find = async (uid: string): Promise<CalendarEventRow[]> => {
        const key = boundIndexedValue(uid);
        try {
            const rows: CalendarEventRow[] = await context.calendarEventRepo.find(
                { icalUid: literalQueryValue(key), mailboxUid: context.mailboxUid, limit: MAX_EVENT_COPIES } as any,
                { ignoreACL: true, limit: MAX_EVENT_COPIES },
            );
            return rows.filter((row) => row.icalUid === key);
        } catch {
            return []; // a UID the query layer can't take as a value matches nothing
        }
    };
    const events = await find(icalUid);
    return events.length > 0 || icalUid.toLowerCase() === icalUid ? events : find(icalUid.toLowerCase());
}

/** `timeZone` when `Intl` accepts it, otherwise `"UTC"`. */
function usableTimeZone(timeZone: string | undefined): string {
    try {
        return new Intl.DateTimeFormat("en-US", { timeZone: timeZone || "UTC" }).resolvedOptions().timeZone;
    } catch {
        return "UTC";
    }
}

/** `date`'s calendar date (`YYYY-MM-DD`) in `timeZone`. */
function dateInZone(date: Date, timeZone: string): string {
    return new Intl.DateTimeFormat("en-CA", { timeZone, year: "numeric", month: "2-digit", day: "2-digit" }).format(date);
}

/** How far `timeZone`'s wall clock is ahead of UTC at `date`, in milliseconds (minute precision). */
function zoneOffsetMs(date: Date, timeZone: string): number {
    const parts: Record<string, string> = {};
    const format = new Intl.DateTimeFormat("en-US", { timeZone, hourCycle: "h23", year: "numeric", month: "numeric", day: "numeric", hour: "numeric", minute: "numeric" });
    for (const part of format.formatToParts(date)) {
        parts[part.type] = part.value;
    }
    const wallClock = Date.UTC(Number(parts.year), Number(parts.month) - 1, Number(parts.day), Number(parts.hour), Number(parts.minute));
    return wallClock - Math.floor(date.getTime() / 60000) * 60000;
}

/** The start of `series`' occurrence on `instanceDate`: the series' local start time on that date in the series'
 * own time zone, so a daylight-saving change in between keeps the same wall-clock time. */
function occurrenceStart(series: CalendarEventRow, instanceDate: string): Date {
    const timeZone = usableTimeZone(series.timezone);
    const start = new Date(series.startDate);
    const days = Math.round((Date.parse(instanceDate) - Date.parse(dateInZone(start, timeZone))) / 86400000);
    const shifted = new Date(start.getTime() + days * 86400000);
    return new Date(shifted.getTime() - (zoneOffsetMs(shifted, timeZone) - zoneOffsetMs(start, timeZone)));
}

/**
 * The copy a response applies to. For an `instanceDate`, the caller's exception copy for that date (its
 * `recurrenceId`'s date in the event's time zone) when one exists; otherwise the series (no `recurrenceId`), with
 * `occurrence` set to the named occurrence's start when the series recurs.
 */
function pickOccurrence(events: CalendarEventRow[], instanceDate: string | undefined): { event: CalendarEventRow; occurrence?: Date } | undefined {
    const exception = instanceDate
        ? events.find((event) => event.recurrenceId != null && dateInZone(new Date(event.recurrenceId), usableTimeZone(event.timezone)) === instanceDate)
        : undefined;
    const series = events.find((event) => event.recurrenceId == null);
    if (exception || !series) {
        return exception ? { event: exception } : undefined;
    }
    return { event: series, occurrence: instanceDate && series.recurrenceRule ? occurrenceStart(series, instanceDate) : undefined };
}

/** Sends the iTIP `REPLY` for `attendee`'s response to the organizer, the same message `BaseCalendarEventRoute.respond()`
 * sends. Throws when the transport doesn't accept it. */
async function sendReply(context: RopContext, event: CalendarEventRow, attendee: Attendee, displayName: string | undefined): Promise<void> {
    const organizerAddress = event.organizer?.address ?? "";
    if (!isPlainEmailAddress(organizerAddress) || !isPlainEmailAddress(attendee.address)) {
        return;
    }
    const ics = buildEventIcs({ ...event, attendees: [attendee] }, "REPLY", { onlyAttendee: attendee });
    const raw: Buffer = await new MailComposer({
        from: { name: displayName ?? "", address: attendee.address },
        to: organizerAddress,
        subject: `${REPLY_SUBJECT_PREFIX[attendee.responseStatus]}: ${event.title}`,
        text: `${attendee.displayName ?? attendee.address} has responded ${attendee.responseStatus} to: ${event.title}`,
        icalEvent: { method: "reply", content: ics },
    })
        .compile()
        .build();
    await sendOrThrow(context.mailTransport, { raw, envelopeFrom: attendee.address, envelopeTo: [organizerAddress] });
}
