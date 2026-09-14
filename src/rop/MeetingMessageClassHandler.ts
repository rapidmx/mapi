///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import MailComposer from "nodemailer/lib/mail-composer/index.js";
import { AttendeeResponseStatus, buildEventIcs, type Attendee, type CalendarEvent } from "@rapidmx/restapi";
import { BufferReader } from "../codec/BufferCursor.js";
import { isPlainEmailAddress } from "./AddressList.js";
import { decodeGlobalObjectId } from "../codec/GlobalObjectId.js";
import type { MapiSessionContext } from "../MapiSessionManager.js";
import { LID_GLOBAL_OBJECT_ID, PSETID_MEETING } from "./CalendarNamedProperties.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import type { RopContext } from "./RopHandler.js";

/** `PidTagMessageClass` suffix -> the caller's own new `AttendeeResponseStatus`, per `[MS-OXOCAL]`'s meeting
 * response object naming convention. An unrecognized suffix (a message class this pragmatic subset doesn't
 * know how to interpret as a response) is simply not handled - see `submitMeetingResponse`'s own doc comment. */
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

/**
 * Handles a submitted `"IPM.Schedule.Meeting.Resp.{Pos,Neg,Tent}"` message - an attendee's own response to a
 * meeting invite this server previously sent (`RopSubmitMessageHandler.submitAppointment`). Delegated to from
 * `RopSubmitMessageHandler` for that message-class prefix, instead of the ordinary mail or Appointment paths.
 *
 * Mirrors restapi's `BaseCalendarEventRoute.respond()`. It correlates the response via `PidLidGlobalObjectId`
 * (echoed by a real client from the invite it's responding to - see `GlobalObjectId.ts`) to **the caller's own
 * copy** of the meeting: the query is scoped to the caller's `mailboxUid`, which is the access check. The
 * organizer's copy (in another mailbox) is never touched here; the organizer's server applies the iTIP `REPLY` sent
 * below, exactly as for a REST response. The earlier version searched every mailbox by `icalUid` and edited
 * whichever event came back first, but an `icalUid` is in every invite each attendee received, so it isn't a secret.
 *
 * For a response to one occurrence of a recurring meeting (the `GlobalObjectId`'s instance date is set), the
 * caller's exception copy for that date is preferred, falling back to the series. The caller's
 * `Attendee.responseStatus` is recorded; declining soft-deletes the caller's copy instead. Then an iTIP `REPLY`
 * (restapi's own `buildEventIcs`) goes to the organizer. A send failure is swallowed, like REST: the response is
 * already recorded.
 *
 * Every failure mode here (no `PidLidGlobalObjectId` set, no matching `CalendarEvent`, no resolvable mailbox,
 * caller isn't actually an attendee, unrecognized message-class suffix) is a silent no-op rather than an error
 * response - the same "don't fail the whole ROP over a stale/unresolvable reference" principle this pragmatic
 * subset applies throughout. A response that can't be correlated or applied is simply dropped.
 */
export async function submitMeetingResponse(messageClass: string, draftProperties: Record<string, string>, context: RopContext): Promise<void> {
    const responseStatus = RESPONSE_STATUS_BY_MESSAGE_CLASS[messageClass];
    if (!responseStatus) {
        return;
    }

    const globalObjectIdBase64 = findNamedPropertyValue(context.session, draftProperties, PSETID_MEETING, LID_GLOBAL_OBJECT_ID);
    if (!globalObjectIdBase64) {
        return;
    }
    const globalObjectId = Buffer.from(globalObjectIdBase64, "base64");
    const icalUid = decodeGlobalObjectId(new BufferReader(globalObjectId));

    const events: CalendarEventRow[] = await context.calendarEventRepo.find(
        { icalUid, mailboxUid: context.mailboxUid, limit: MAX_EVENT_COPIES },
        { ignoreACL: true, limit: MAX_EVENT_COPIES },
    );
    const event = pickOccurrence(events, instanceDateOf(globalObjectId));
    if (!event) {
        return;
    }

    const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
    if (!mailbox) {
        return;
    }
    const callerAddresses = new Set([mailbox.primarySmtpAddress.toLowerCase(), ...mailbox.aliasAddresses.map((a: string) => a.toLowerCase())]);
    const respondingAttendee = event.attendees.find((attendee) => callerAddresses.has(attendee.address.toLowerCase()));
    if (!respondingAttendee) {
        return;
    }

    const updatedAttendee: Attendee = { ...respondingAttendee, responseStatus };
    if (responseStatus === AttendeeResponseStatus.DECLINED) {
        await context.calendarEventRepo.delete(event.uid, { ignoreACL: true });
    } else {
        const attendees: Attendee[] = event.attendees.map((attendee) => (attendee === respondingAttendee ? updatedAttendee : attendee));
        await context.calendarEventRepo.update({ uid: event.uid, version: event.version, attendees }, event, { ignoreACL: true });
    }

    await sendReply(context, event, updatedAttendee, mailbox.displayName);
}

/** The most copies of one meeting (series plus exceptions) considered in the caller's mailbox. */
const MAX_EVENT_COPIES = 100;

type CalendarEventRow = CalendarEvent & { uid: string; version: number };

const REPLY_SUBJECT_PREFIX: Record<string, string> = {
    [AttendeeResponseStatus.ACCEPTED]: "Accepted",
    [AttendeeResponseStatus.TENTATIVE]: "Tentative",
    [AttendeeResponseStatus.DECLINED]: "Declined",
};

/** The `YYYY-MM-DD` (UTC) occurrence a `GlobalObjectId` names (`[MS-OXOCAL]` `YH`/`YL`/`M`/`D` bytes), or
 * `undefined` for the whole series (all zero). */
function instanceDateOf(globalObjectId: Buffer): string | undefined {
    const year = (globalObjectId[16] << 8) | globalObjectId[17];
    if (year === 0) {
        return undefined;
    }
    return `${String(year).padStart(4, "0")}-${String(globalObjectId[18]).padStart(2, "0")}-${String(globalObjectId[19]).padStart(2, "0")}`;
}

/** The caller's copy for `instanceDate` when one exists, otherwise the series itself (no `recurrenceId`). */
function pickOccurrence(events: CalendarEventRow[], instanceDate: string | undefined): CalendarEventRow | undefined {
    const exception = instanceDate
        ? events.find((event) => event.recurrenceId != null && new Date(event.recurrenceId).toISOString().slice(0, 10) === instanceDate)
        : undefined;
    return exception ?? events.find((event) => event.recurrenceId == null);
}

/** Sends the iTIP `REPLY` for `attendee`'s response to the organizer, the same message `BaseCalendarEventRoute.respond()`
 * sends. Best-effort: a failure is swallowed, since the response itself is already saved. */
async function sendReply(context: RopContext, event: CalendarEventRow, attendee: Attendee, displayName: string | undefined): Promise<void> {
    const organizerAddress = event.organizer?.address ?? "";
    if (!isPlainEmailAddress(organizerAddress) || !isPlainEmailAddress(attendee.address)) {
        return;
    }
    try {
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
        await context.mailTransport.send({ raw, envelopeFrom: attendee.address, envelopeTo: [organizerAddress] });
    } catch {
        // Same as REST: the response is already recorded, so a failed notification doesn't fail the request.
    }
}
