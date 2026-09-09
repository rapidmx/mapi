///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { AttendeeResponseStatus, type Attendee, type CalendarEvent } from "@rapidmx/restapi";
import { BufferReader } from "../codec/BufferCursor.js";
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
 * Correlates the response back to the original `CalendarEvent` via `PidLidGlobalObjectId` (echoed by a real
 * client from the invite it's responding to - see `GlobalObjectId.ts`'s own doc comment for why this requires
 * no separate correlation table) and updates the caller's own `Attendee.responseStatus` in place, reusing the
 * exact same matching-by-mailbox-address logic `MeetingResponseCommand` (EAS's own `MeetingResponse` handler)
 * already implements - this is a from-scratch reimplementation of that small algorithm rather than a
 * cross-import, since EAS and MAPI are independent protocol verticals in this codebase with no dependency
 * between them.
 *
 * Every failure mode here (no `PidLidGlobalObjectId` set, no matching `CalendarEvent`, no resolvable mailbox,
 * caller isn't actually an attendee, unrecognized message-class suffix) is a silent no-op rather than an error
 * response - the same "don't fail the whole ROP over a stale/unresolvable reference" principle this pragmatic
 * subset applies throughout (e.g. `MessageTarget.resolveMessageInfo`'s own doc comment). A response that can't
 * be correlated or applied is simply dropped, not a reason to reject the client's `RopSubmitMessage` call.
 *
 * **The one query in this package that isn't `mailboxUid`-scoped, and why that's unavoidable here**: unlike
 * every other `ignoreACL: true` repo call in this codebase (all transitively scoped to the caller's own
 * mailbox via a session FID/MID registry that itself only ever gets populated from mailbox-scoped queries),
 * `calendarEventRepo.find({ icalUid })` below deliberately searches *every* mailbox on this server - the
 * `CalendarEvent` being responded to belongs to the meeting's *organizer*, who is very often a different
 * mailbox than the attendee submitting this response, and the caller's own session has no way to know which
 * mailbox that is. Safety instead comes from two independent properties: `icalUid` is a server-generated
 * `crypto.randomUUID()` (see `RopSaveChangesMessageHandler.saveAppointment`), for all practical purposes
 * globally unique, so this can't be used to enumerate or collide with another organizer's events; and the
 * actual authorization gate is the attendee-membership check a few lines down (`callerAddresses.has(...)`),
 * derived entirely from *this* caller's own resolved mailbox record, not from anything the client supplied -
 * a crafted `PidLidGlobalObjectId` can at best name a real `icalUid` it doesn't already know the value of
 * (astronomically unlikely against a random UUID) and still can't mutate an event unless the caller's own
 * mailbox is genuinely a listed attendee on it.
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
    const icalUid = decodeGlobalObjectId(new BufferReader(Buffer.from(globalObjectIdBase64, "base64")));

    const events: (CalendarEvent & { uid: string; version: number })[] = await context.calendarEventRepo.find(
        { icalUid },
        { ignoreACL: true, limit: 1 },
    );
    const event = events[0];
    if (!event) {
        return;
    }

    const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
    if (!mailbox) {
        return;
    }
    const callerAddresses = new Set([mailbox.primarySmtpAddress.toLowerCase(), ...mailbox.aliasAddresses.map((a: string) => a.toLowerCase())]);
    const attendeeIndex = event.attendees.findIndex((attendee) => callerAddresses.has(attendee.address.toLowerCase()));
    if (attendeeIndex === -1) {
        return;
    }

    const attendees: Attendee[] = event.attendees.map((attendee, i) => (i === attendeeIndex ? { ...attendee, responseStatus } : attendee));
    await context.calendarEventRepo.update({ uid: event.uid, version: event.version, attendees }, event, { ignoreACL: true });
}
