///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import * as crypto from "crypto";
import MailComposer from "nodemailer/lib/mail-composer/index.js";
import {
    findOrCreateWellKnownFolder,
    FolderType,
    MessageImportance,
    RecipientType,
    scanAndRelay,
    type CalendarEvent,
} from "@rapidmx/restapi";
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { MapiObjectHandle } from "../MapiSessionManager.js";
import { isPlainEmailAddress, splitAddressList } from "./AddressList.js";
import { submitMeetingResponse } from "./MeetingMessageClassHandler.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_SUBMIT_MESSAGE = 0x32;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for both "the referenced handle isn't a draft
 * message" and "the draft has no resolvable recipients/sender" - the same constant reused throughout this
 * pragmatic subset as a generic "can't do this" signal, not a claim of exact per-condition `[MS-OXCRPC]`
 * return-value-table parity. */
const ERROR_INVALID_OBJECT = 0x80070005;

// The same well-known property IDs RopSetPropertiesHandler tracks - duplicated here (rather than imported)
// since importing them would only save a handful of literals; see that file for the real documentation of what
// each means.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_CLASS = 0x001a;
const PID_TAG_DISPLAY_BCC = 0x0e02;
const PID_TAG_DISPLAY_CC = 0x0e03;
const PID_TAG_DISPLAY_TO = 0x0e04;
const PID_TAG_BODY = 0x1000;
const PID_TAG_READ_RECEIPT_REQUESTED = 0x0029;
const PID_TAG_DEFERRED_SEND_TIME = 0x3fef;

/** `PidTagMessageClass` values starting with this prefix (`IPM.Appointment`, `IPM.Appointment.*`) route through
 * `submitAppointment()` instead of the ordinary mail compose/send path below - see that method's own doc
 * comment. */
const MESSAGE_CLASS_APPOINTMENT_PREFIX = "IPM.Appointment";

/** `PidTagMessageClass` values starting with this prefix (`IPM.Schedule.Meeting.Resp.{Pos,Neg,Tent}`) route
 * through `MeetingMessageClassHandler.submitMeetingResponse()` instead of every other path here - see that
 * function's own doc comment. */
const MESSAGE_CLASS_MEETING_RESPONSE_PREFIX = "IPM.Schedule.Meeting.Resp.";

// Re-exported for existing importers; the implementation lives in AddressList.ts.
export { isPlainEmailAddress, parseAddressList } from "./AddressList.js";

/**
 * `RopSubmitMessage` (`[MS-OXOMSG]`/`[MS-OXCROPS]`): sends a composed message - the actual trigger point for
 * this pragmatic subset's compose/send pipeline, since `RopSaveChangesMessage` itself does no real persistence
 * (see its own doc comment). Builds a raw MIME buffer from whatever `RopSetProperties`/`RopWriteStream` calls
 * accumulated on the draft handle (`Subject`, the `PidTagBody` stream or an inline `PidTagBody` property, and
 * addressing - see below), via `nodemailer`'s `MailComposer` (already a dependency, used elsewhere in this
 * library only to *send*, not build, MIME - this is its first use purely as a RFC 5322 serializer), then calls
 * the exact same `scanAndRelay()` scan-then-relay pipeline `BaseMessageRoute.send()`/EAS's `ComposeMailCommand`
 * already share, and persists a Sent Items copy (mirroring `ComposeMailCommand`'s own always-committed Sent
 * Items behavior for MAPI submission specifically, since - unlike EAS's optional `SaveInSentItems` - a MAPI
 * `RopSubmitMessage` call has no "don't keep a copy" option to honor).
 *
 * **Known, deliberate limitation - no `RopModifyRecipients` support**: real recipient rows (`[MS-OXCMSG]`
 * §2.2.3.2, each with its own `PidTagEmailAddress`/`PidTagRecipientType`) are a substantial structure this
 * pragmatic subset's ROP coverage doesn't include. Instead, addressing is read from `PidTagDisplayTo`/
 * `DisplayCc`/`DisplayBcc` - the same semicolon-separated "cached recipient display string" properties real
 * Outlook *also* always sets via `RopSetProperties` alongside `RopModifyRecipients` (as a display-only
 * convenience cache). This works correctly for the common case of composing to bare, unresolved email
 * addresses (nothing to resolve to a distinct display name) - the case this library's own integration tests
 * exercise - but is **not** a substitute for real recipient rows: a client that resolves typed names against an
 * address book before submitting would populate these fields with display names, not addresses, and this
 * pragmatic subset has no way to recover a real SMTP address from a bare display name. Documented here rather
 * than silently producing wrong addressing.
 *
 * `SubmitFlags` (`PreprocessOnly`, ...) is decoded to advance past it correctly but not honored - this
 * pragmatic subset has no transport-agent preprocessing distinction to vary by flag.
 *
 * **`PidTagDeferredSendTime`** ("Do not deliver before"), if set to a future time via `RopSetProperties`,
 * parks the message in Outbox with `Message.scheduledSendTime` set instead of relaying immediately - mirroring
 * `BaseMessageRoute.send()`'s own identical branch - for `ScheduledSendJob` to relay later via this same
 * `scanAndRelay()` pipeline. **`PidTagReadReceiptRequested`**, if `true`, attaches a real
 * `Disposition-Notification-To` header (see `util/ReceiptUtils.ts`) and is recorded on the persisted message's
 * own `requestReceipt` field - honored only when the client explicitly sets it, not this mailbox's
 * `alwaysRequestReceiptInternal`/`External` defaults (those apply to the REST/webmail compose path only).
 *
 * **Calendar branches**: a draft whose `PidTagMessageClass` starts with `"IPM.Appointment"` is routed to
 * `submitAppointment()` instead of the mail path below - see that method's own doc comment. One starting with
 * `"IPM.Schedule.Meeting.Resp."` (an attendee's own accept/decline/tentative response to a meeting this server
 * previously invited them to) is routed to `MeetingMessageClassHandler.submitMeetingResponse()` instead - see
 * that function's own doc comment.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSubmitMessageHandler implements RopHandler {
    public readonly ropId = ROP_ID_SUBMIT_MESSAGE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // SubmitFlags - see class doc comment

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "message") {
            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const properties = handle.draftProperties ?? {};
        const messageClass = properties[String(PID_TAG_MESSAGE_CLASS)] ?? "IPM.Note";
        if (messageClass.startsWith(MESSAGE_CLASS_APPOINTMENT_PREFIX)) {
            await this.submitAppointment(handle, context, writer, inputHandleIndex);
            return;
        }
        if (messageClass.startsWith(MESSAGE_CLASS_MEETING_RESPONSE_PREFIX)) {
            await submitMeetingResponse(messageClass, properties, context);
            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(0); // ReturnValue - success, see MeetingMessageClassHandler's own doc comment
            // on why every failure mode there is a silent no-op rather than surfaced here.
            return;
        }

        const to = splitAddressList(properties[String(PID_TAG_DISPLAY_TO)]);
        const cc = splitAddressList(properties[String(PID_TAG_DISPLAY_CC)]);
        const bcc = splitAddressList(properties[String(PID_TAG_DISPLAY_BCC)]);
        const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
        const envelopeFrom: string | undefined = mailbox?.primarySmtpAddress;

        // Refuse the whole send over one malformed recipient rather than silently dropping it.
        const allValid = [...to, ...cc, ...bcc].every(isPlainEmailAddress);
        if (!envelopeFrom || to.length + cc.length + bcc.length === 0 || !allValid) {
            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const subject = properties[String(PID_TAG_SUBJECT)] ?? "";
        const bodyText = resolveDraftBody(context, inputHandleIndex, properties);
        const requestReceipt = properties[String(PID_TAG_READ_RECEIPT_REQUESTED)] === "true";
        // "Do not deliver before" - a real Outlook compose option, tracked as a plain ISO-8601 string by
        // RopSetPropertiesHandler.stringifyValue()'s PtypTime encoding.
        const deferredSendTimeRaw = properties[String(PID_TAG_DEFERRED_SEND_TIME)];
        const deferredSendTime: Date | undefined = deferredSendTimeRaw ? new Date(deferredSendTimeRaw) : undefined;
        const isDeferred = deferredSendTime !== undefined && deferredSendTime.getTime() > Date.now();

        const raw: Buffer = await new MailComposer({
            from: envelopeFrom,
            to,
            cc,
            bcc,
            subject,
            text: bodyText,
            // A real Disposition-Notification-To request, mirroring BaseMessageRoute.send()'s own explicit
            // Message.requestReceipt handling (see util/ReceiptUtils.ts) - honored here only when the client
            // itself asked for one via PidTagReadReceiptRequested, not this mailbox's own
            // alwaysRequestReceiptInternal/External defaults (those apply to the REST/webmail compose path,
            // which already goes through BaseMessageRoute.send() directly).
            ...(requestReceipt ? { headers: [{ key: "Disposition-Notification-To", value: envelopeFrom }] } : {}),
        })
            .compile()
            .build();
        const envelopeTo = [...to, ...cc, ...bcc];

        const bodyBlobKey = `bodies/${crypto.randomUUID()}`;
        const recipients = [
            ...to.map((address) => ({ address, type: RecipientType.TO })),
            ...cc.map((address) => ({ address, type: RecipientType.CC })),
            ...bcc.map((address) => ({ address, type: RecipientType.BCC })),
        ];

        if (isDeferred) {
            // Mirrors BaseMessageRoute.send()'s own scheduledSendTime branch: park the message in Outbox,
            // unscanned/unrelayed, for the same ScheduledSendJob to pick up and relay (via this exact
            // scanAndRelay() pipeline) once due - see that job's own doc comment.
            await context.blobStore.put(bodyBlobKey, raw, { contentType: "message/rfc822" });
            const outbox = await findOrCreateWellKnownFolder(context.folderRepo, context.folderClass, context.mailboxUid, FolderType.OUTBOX);
            await context.messageRepo.create(
                new context.messageClass({
                    folderUid: outbox.uid,
                    mailboxUid: context.mailboxUid,
                    messageId: `${crypto.randomUUID()}@mapi`,
                    subject,
                    from: { address: envelopeFrom, type: RecipientType.TO },
                    recipients,
                    sentDate: new Date(),
                    receivedDate: new Date(),
                    bodyBlobKey,
                    bodyPreview: bodyText.slice(0, 200),
                    flags: { read: true, flagged: false, answered: false, forwarded: false },
                    importance: MessageImportance.NORMAL,
                    references: [],
                    hasAttachments: false,
                    scheduledSendTime: deferredSendTime,
                    requestReceipt,
                }),
                { ignoreACL: true },
            );

            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(0); // ReturnValue - success
            return;
        }

        const relayed = await scanAndRelay(raw, envelopeFrom, envelopeTo, context.scanPipeline, context.mailTransport, context.blobStore);

        // The Sent Items copy stores exactly what was relayed: scanAndRelay may have added a Message-ID header, and
        // that header's value (plus the thread it derived) is what recipients' copies carry and what recall targets.
        await context.blobStore.put(bodyBlobKey, relayed.raw, { contentType: "message/rfc822" });
        const sentFolder = await findOrCreateWellKnownFolder(context.folderRepo, context.folderClass, context.mailboxUid, FolderType.SENT_ITEMS);
        await context.messageRepo.create(
            new context.messageClass({
                folderUid: sentFolder.uid,
                mailboxUid: context.mailboxUid,
                messageId: relayed.messageId,
                conversationId: relayed.conversationId,
                encrypted: relayed.encrypted,
                subject,
                from: { address: envelopeFrom, type: RecipientType.TO },
                recipients,
                sentDate: new Date(),
                receivedDate: new Date(),
                bodyBlobKey,
                sanitizedHtmlBlobKey: relayed.sanitizedHtmlBlobKey,
                bodyPreview: bodyText.slice(0, 200),
                flags: { read: true, flagged: false, answered: false, forwarded: false },
                importance: MessageImportance.NORMAL,
                references: [],
                hasAttachments: false,
                requestReceipt,
            }),
            { ignoreACL: true },
        );

        writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }

    /**
     * Submits a Calendar item. The appointment itself is already persisted (`RopSaveChangesMessageHandler`
     * writes the real `CalendarEvent` row at Save time, unlike a mail draft, which stays purely in-session until
     * Submit) - this method's only job is notifying attendees, if any were set (via `PidTagDisplayTo`/`Cc`,
     * decoded into `CalendarEvent.attendees` at Save time). No attendees means no invite to send (e.g. a private,
     * non-meeting appointment) - a well-formed success response either way, since the appointment already exists.
     *
     * Unlike ordinary mail, there is no separate "Sent Items copy" to create - the appointment's durable copy
     * *is* the `CalendarEvent` row already sitting in the organizer's own Calendar folder.
     *
     * The invite itself is a minimal iCalendar (`RFC 5545`) `VEVENT` with `METHOD:REQUEST` (`RFC 5546`),
     * attached via `nodemailer`'s own `MailComposer` `icalEvent` option (which produces both a `text/calendar;
     * method=REQUEST` MIME alternative and an `.ics` attachment - the standard dual form real invite emails use)
     * - reusing the identical `scanAndRelay()` pipeline the mail path above already calls.
     */
    private async submitAppointment(handle: MapiObjectHandle, context: RopContext, writer: BufferWriter, inputHandleIndex: number): Promise<void> {
        const uid = handle.entityUid.startsWith("calendarEvent:") ? handle.entityUid.slice("calendarEvent:".length) : undefined;
        const event: CalendarEvent | undefined = uid ? await context.calendarEventRepo.findOne(uid, { ignoreACL: true }) : undefined;
        const mailbox = event?.mailboxUid === context.mailboxUid ? await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true }) : undefined;
        const callerAddresses: string[] = mailbox ? [mailbox.primarySmtpAddress, ...(mailbox.aliasAddresses ?? [])].map((a: string) => a.toLowerCase()) : [];
        // Only the organizer sends invites. An attendee's own copy of a meeting (delivered into their calendar) is
        // also a CalendarEvent they can open and submit, which would otherwise mail a REQUEST to every attendee as if
        // they had organized it.
        if (!event || !callerAddresses.includes((event.organizer?.address ?? "").toLowerCase())) {
            writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const attendees = event.attendees.filter((attendee) => isPlainEmailAddress(attendee.address));
        if (attendees.length > 0) {
            const envelopeFrom: string = mailbox.primarySmtpAddress;
            const attendeeAddresses = attendees.map((attendee) => attendee.address);
            const ics = buildMeetingRequestIcs({ ...event, attendees }, envelopeFrom);
            const raw: Buffer = await new MailComposer({
                from: envelopeFrom,
                to: attendeeAddresses,
                subject: event.title,
                text: `You have been invited to: ${event.title}`,
                icalEvent: { method: "REQUEST", content: ics },
            })
                .compile()
                .build();
            await scanAndRelay(raw, envelopeFrom, attendeeAddresses, context.scanPipeline, context.mailTransport, context.blobStore);
        }

        writer.writeUInt8(ROP_ID_SUBMIT_MESSAGE);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }
}

/** `YYYYMMDDTHHMMSSZ` - the `RFC 5545` "form 2" (UTC) `DATE-TIME` format every field below uses. */
function formatIcsDateUtc(date: Date): string {
    return date.toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
}

/** Escapes the `RFC 5545` §3.3.11 `TEXT` value special characters (backslash, semicolon, comma, newline) - the
 * only value types this minimal builder ever emits unescaped text into (`SUMMARY`/`LOCATION`). */
function escapeIcsText(value: string): string {
    return value.replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\r?\n/g, "\\n");
}

/**
 * Builds a minimal `RFC 5545`/`RFC 5546` `METHOD:REQUEST` iCalendar document for `event`. Deliberately not a
 * general-purpose iCalendar writer (see `AutodiscoverXml.ts`'s own doc comment for the same "hand-build just
 * the fields this pass needs" precedent) - no `RRULE` (recurring meeting invites are a documented gap, same as
 * this pragmatic subset's other recurrence-adjacent limitations), no line-folding at the 75-octet boundary
 * `RFC 5545` §3.1 technically requires (every field emitted here is short enough in practice that folding would
 * never trigger), no `VTIMEZONE` block (times are always emitted as UTC `Z`-suffixed values instead).
 */
function buildMeetingRequestIcs(event: CalendarEvent, organizerAddress: string): string {
    const lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//RapidREST//Mail//EN",
        "VERSION:2.0",
        "METHOD:REQUEST",
        "BEGIN:VEVENT",
        `UID:${event.icalUid}`,
        `SEQUENCE:${event.sequence}`,
        `DTSTAMP:${formatIcsDateUtc(new Date())}`,
        `DTSTART:${formatIcsDateUtc(event.startDate)}`,
        `DTEND:${formatIcsDateUtc(event.endDate)}`,
        `SUMMARY:${escapeIcsText(event.title)}`,
        ...(event.location ? [`LOCATION:${escapeIcsText(event.location)}`] : []),
        `ORGANIZER:mailto:${organizerAddress}`,
        ...event.attendees.map((attendee) => `ATTENDEE;RSVP=TRUE:mailto:${attendee.address}`),
        "STATUS:CONFIRMED",
        "END:VEVENT",
        "END:VCALENDAR",
    ];
    return lines.join("\r\n");
}

/** Prefers a `RopWriteStream`-accumulated body (the real path a large body takes) over an inline `PidTagBody`
 * set directly via `RopSetProperties` (only realistic for a short body small enough to set inline) - see
 * `RopOpenStreamHandler`'s write-mode branch for how `writeTargetHandleIndex` links a stream handle back to the
 * draft message handle it was opened against. The generation must match too, so a stream left over from an
 * earlier message at the same handle index never supplies this message's body. */
function resolveDraftBody(context: RopContext, messageHandleIndex: number, properties: Record<string, string>): string {
    const generation = context.session.handles[messageHandleIndex].generation;
    for (const candidate of Object.values(context.session.handles)) {
        if (
            candidate.type === "stream" &&
            candidate.writeTargetHandleIndex === messageHandleIndex &&
            candidate.writeTargetGeneration === generation &&
            candidate.writeBufferBase64
        ) {
            const raw = Buffer.from(candidate.writeBufferBase64, "base64");
            return raw.toString("utf16le").replace(/\0+$/, "");
        }
    }
    return properties[String(PID_TAG_BODY)] ?? "";
}
