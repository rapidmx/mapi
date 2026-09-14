///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { encodeGlobalObjectId } from "../../src/codec/GlobalObjectId.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { submitMeetingResponse } from "../../src/rop/MeetingMessageClassHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { AttendeeResponseStatus, AttendeeRole } from "@rapidmx/restapi";

const PSETID_MEETING = "6ed8da90-450b-101b-98da-00aa003f1305";
const LID_GLOBAL_OBJECT_ID = 0x00000003;

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: {} as any,
        calendarEventRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
        ...overrides,
    };
}

function globalObjectIdProperty(session: MapiSessionContext, icalUid: string): Record<string, string> {
    const id = assignOrGetNamedPropertyId(session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
    return { [String(id)]: encodeGlobalObjectId(icalUid, new Date()).toString("base64") };
}

describe("submitMeetingResponse Tests", () => {
    it("Returns MAPI_E_INVALID_PARAMETER for an unrecognized message-class suffix.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn() } as any });
        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Bogus", globalObjectIdProperty(context.session, "x"), context)).toBe(0x80070057);
        expect((context.calendarEventRepo as any).find).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_PARAMETER when the draft never set PidLidGlobalObjectId, or set a malformed one.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn() } as any });
        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", {}, context)).toBe(0x80070057);
        const id = assignOrGetNamedPropertyId(context.session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", { [String(id)]: Buffer.alloc(40).toString("base64") }, context)).toBe(0x80070057);
        expect((context.calendarEventRepo as any).find).not.toHaveBeenCalled();
    });

    it("Skips past an unrelated named property to find PidLidGlobalObjectId among several draft properties.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn().mockResolvedValue([]) } as any });
        const unrelatedId = assignOrGetNamedPropertyId(context.session, {
            guid: "00062002-0000-0000-c000-000000000046",
            kind: "lid",
            lid: 0x8208,
        });
        const properties = {
            "26": "IPM.Schedule.Meeting.Resp.Pos", // a plain (< 0x8000) PidTag property mixed in among the named ones
            [String(unrelatedId)]: "unrelated-value",
            ...globalObjectIdProperty(context.session, "evt-uid@example.com"),
        };

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect((context.calendarEventRepo as any).find).toHaveBeenCalledWith(
            { icalUid: "eq(evt-uid@example.com)", mailboxUid: "mailbox-1", limit: 100 },
            { ignoreACL: true, limit: 100 },
        );
    });

    it("Returns MAPI_E_NOT_FOUND when no CalendarEvent matches the decoded icalUid.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn().mockResolvedValue([]) } as any });
        const properties = globalObjectIdProperty(context.session, "unknown@example.com");

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context)).toBe(0x8004010f);

        expect((context.calendarEventRepo as any).find).toHaveBeenCalledWith(
            { icalUid: "eq(unknown@example.com)", mailboxUid: "mailbox-1", limit: 100 },
            { ignoreACL: true, limit: 100 },
        );
    });

    it("Returns MAPI_E_NOT_FOUND when the caller's mailbox can't be resolved.", async () => {
        const event = { uid: "evt1", version: 1, icalUid: "evt-uid@example.com", attendees: [{ address: "caller@example.com" }] };
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update: vi.fn() } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context)).toBe(0x8004010f);

        expect((context.calendarEventRepo as any).update).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_NOT_FOUND when the caller isn't actually an attendee of the event.", async () => {
        const event = { uid: "evt1", version: 1, icalUid: "evt-uid@example.com", attendees: [{ address: "someone-else@example.com" }] };
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update: vi.fn() } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context)).toBe(0x8004010f);

        expect((context.calendarEventRepo as any).update).not.toHaveBeenCalled();
    });

    it("Updates the caller's own Attendee.responseStatus to ACCEPTED for .Resp.Pos, matched by primarySmtpAddress (case-insensitive).", async () => {
        const event = {
            uid: "evt1",
            version: 2,
            icalUid: "evt-uid@example.com",
            attendees: [
                { address: "Caller@Example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.NEEDS_ACTION, isOrganizer: false },
                { address: "other@example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.NEEDS_ACTION, isOrganizer: false },
            ],
        };
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([event]), update: vi.fn().mockResolvedValue(undefined) };
        const context = makeContext({
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect(calendarEventRepo.update).toHaveBeenCalledTimes(1);
        const [delta, existingArg, options] = calendarEventRepo.update.mock.calls[0];
        expect(delta.uid).toBe("evt1");
        expect(delta.version).toBe(2);
        expect(delta.attendees[0]).toEqual({ ...event.attendees[0], responseStatus: AttendeeResponseStatus.ACCEPTED });
        expect(delta.attendees[1]).toEqual(event.attendees[1]); // untouched
        expect(existingArg).toBe(event);
        expect(options).toEqual({ ignoreACL: true });
    });

    it("Updates to TENTATIVE for .Resp.Tent.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
            icalUid: "evt-uid@example.com",
            attendees: [{ address: "caller@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION }],
        };
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([event]), update: vi.fn().mockResolvedValue(undefined) };
        const context = makeContext({
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Tent", properties, context);

        const [delta] = calendarEventRepo.update.mock.calls[0];
        expect(delta.attendees[0].responseStatus).toBe(AttendeeResponseStatus.TENTATIVE);
    });

    it("Declining (.Resp.Neg) soft-deletes the caller's own copy, like REST respond(), instead of updating it.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
            icalUid: "evt-uid@example.com",
            attendees: [{ address: "caller@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION }],
        };
        const calendarEventRepo = {
            find: vi.fn().mockResolvedValue([event]),
            update: vi.fn().mockResolvedValue(undefined),
            delete: vi.fn().mockResolvedValue(undefined),
        };
        const context = makeContext({
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", properties, context);

        expect(calendarEventRepo.delete).toHaveBeenCalledWith("evt1", { ignoreACL: true });
        expect(calendarEventRepo.update).not.toHaveBeenCalled();
    });

    it("Sends an iTIP REPLY carrying only the caller's new status to the organizer.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
            icalUid: "evt-uid@example.com",
            sequence: 3,
            status: "confirmed",
            title: "Planning",
            startDate: new Date("2026-10-01T10:00:00.000Z"),
            endDate: new Date("2026-10-01T11:00:00.000Z"),
            organizer: { address: "boss@example.com" },
            attendees: [
                { address: "caller@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION },
                { address: "other@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION },
            ],
        };
        const send = vi.fn().mockResolvedValue({ accepted: ["boss@example.com"], rejected: [] });
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [], displayName: "Caller" }) } as any,
            mailTransport: { send } as any,
        });

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Tent", globalObjectIdProperty(context.session, "evt-uid@example.com"), context)).toBe(0);

        expect(send).toHaveBeenCalledTimes(1);
        const { raw, envelopeFrom, envelopeTo } = send.mock.calls[0][0];
        expect(envelopeFrom).toBe("caller@example.com");
        expect(envelopeTo).toEqual(["boss@example.com"]);
        const mime = raw.toString();
        expect(mime).toContain("Subject: Tentative: Planning");
        expect(mime).toMatch(/method=(REPLY|reply)/);
        expect(mime).not.toContain("other@example.com");
    });

    it("Records the response but reports MAPI_E_CALL_FAILED when the REPLY is rejected or fails, and skips a malformed organizer address.", async () => {
        const makeEvent = (organizer: string) => ({
            uid: "evt1",
            version: 1,
            icalUid: "evt-uid@example.com",
            sequence: 0,
            status: "confirmed",
            title: "T",
            startDate: new Date(),
            endDate: new Date(),
            organizer: { address: organizer },
            attendees: [{ address: "caller@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION }],
        });
        const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) };

        const failingSend = vi.fn().mockRejectedValue(new Error("relay down"));
        const failing = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([makeEvent("boss@example.com")]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: mailboxRepo as any,
            mailTransport: { send: failingSend } as any,
        });
        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(failing.session, "evt-uid@example.com"), failing)).toBe(0x80004005);
        expect((failing.calendarEventRepo as any).update).toHaveBeenCalledTimes(1);
        expect(failingSend).toHaveBeenCalledTimes(1);

        // A transport reporting a rejected recipient, instead of throwing, is a failure too (sendOrThrow).
        for (const result of [{ accepted: [], rejected: ["boss@example.com"] }, { accepted: ["boss@example.com"], rejected: ["x@example.com"] }, {}, undefined]) {
            const rejecting = makeContext({
                calendarEventRepo: { find: vi.fn().mockResolvedValue([makeEvent("boss@example.com")]), update: vi.fn().mockResolvedValue(undefined) } as any,
                mailboxRepo: mailboxRepo as any,
                mailTransport: { send: vi.fn().mockResolvedValue(result) } as any,
            });
            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(rejecting.session, "evt-uid@example.com"), rejecting)).toBe(0x80004005);
        }

        const injectedSend = vi.fn();
        const injected = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([makeEvent("boss@example.com\r\nBcc: victim@example.com")]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: mailboxRepo as any,
            mailTransport: { send: injectedSend } as any,
        });
        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(injected.session, "evt-uid@example.com"), injected)).toBe(0);
        expect(injectedSend).not.toHaveBeenCalled();
    });

    it("Prefers the caller's exception copy for the occurrence the GlobalObjectId names, falling back to the series.", async () => {
        const series = { uid: "series", version: 1, icalUid: "evt-uid@example.com", attendees: [{ address: "caller@example.com" }] };
        const exception = { uid: "exception", version: 1, icalUid: "evt-uid@example.com", recurrenceId: new Date("2026-10-08T10:00:00.000Z"), attendees: [{ address: "caller@example.com" }] };
        const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) };
        const instanceProperty = (session: MapiSessionContext, year: number, month: number, day: number): Record<string, string> => {
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
            const bytes = encodeGlobalObjectId("evt-uid@example.com", new Date());
            bytes[16] = year >> 8;
            bytes[17] = year & 0xff;
            bytes[18] = month;
            bytes[19] = day;
            return { [String(id)]: bytes.toString("base64") };
        };

        const onException = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([series, exception]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: mailboxRepo as any,
        });
        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", instanceProperty(onException.session, 2026, 10, 8), onException);
        expect((onException.calendarEventRepo as any).update.mock.calls[0][0].uid).toBe("exception");

        const noException = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([exception, series]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: mailboxRepo as any,
        });
        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", instanceProperty(noException.session, 2026, 10, 15), noException);
        expect((noException.calendarEventRepo as any).update.mock.calls[0][0].uid).toBe("series");

        const seriesResponse = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([exception, series]), update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: mailboxRepo as any,
        });
        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(seriesResponse.session, "evt-uid@example.com"), seriesResponse);
        expect((seriesResponse.calendarEventRepo as any).update.mock.calls[0][0].uid).toBe("series");
    });

    it("Returns MAPI_E_NOT_FOUND when the caller holds only exception copies and none matches.", async () => {
        const exception = { uid: "exception", version: 1, icalUid: "evt-uid@example.com", recurrenceId: new Date("2026-10-08T10:00:00.000Z"), attendees: [{ address: "caller@example.com" }] };
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([exception]), update: vi.fn() } as any,
            mailboxRepo: { findOne: vi.fn() } as any,
        });

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(context.session, "evt-uid@example.com"), context)).toBe(0x8004010f);

        expect((context.calendarEventRepo as any).update).not.toHaveBeenCalled();
        expect((context.mailboxRepo as any).findOne).not.toHaveBeenCalled();
    });


    describe("One occurrence of a recurring series", () => {
        const mailboxRepo = () => ({ findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [], displayName: "Caller" }) });
        const makeSeries = (overrides: Record<string, unknown> = {}) => ({
            uid: "series",
            version: 4,
            icalUid: "evt-uid@example.com",
            sequence: 1,
            status: "confirmed",
            title: "Weekly",
            timezone: "America/Los_Angeles",
            // 09:00 Pacific daylight time on Thursday 2026-10-01.
            startDate: new Date("2026-10-01T16:00:00.000Z"),
            endDate: new Date("2026-10-01T16:30:00.000Z"),
            recurrenceRule: { freq: "weekly", interval: 1, byDay: ["TH"], exceptions: [new Date("2026-10-08T16:00:00.000Z")] },
            organizer: { address: "boss@example.com" },
            attendees: [{ address: "caller@example.com", responseStatus: AttendeeResponseStatus.ACCEPTED }],
            ...overrides,
        });
        const instanceProperty = (session: MapiSessionContext, year: number, month: number, day: number): Record<string, string> => {
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
            const bytes = encodeGlobalObjectId("evt-uid@example.com", new Date());
            bytes[16] = year >> 8;
            bytes[17] = year & 0xff;
            bytes[18] = month;
            bytes[19] = day;
            return { [String(id)]: bytes.toString("base64") };
        };

        it("Declining one occurrence adds it to the caller's series exceptions instead of deleting the series, and replies with RECURRENCE-ID.", async () => {
            const series = makeSeries();
            const send = vi.fn().mockResolvedValue({ accepted: ["boss@example.com"], rejected: [] });
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([series]), update: vi.fn().mockResolvedValue(undefined), delete: vi.fn() };
            const context = makeContext({ calendarEventRepo: calendarEventRepo as any, mailboxRepo: mailboxRepo() as any, mailTransport: { send } as any });

            // 2026-11-05 is after the switch back to standard time: the occurrence is still 09:00 local, now 17:00 UTC.
            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", instanceProperty(context.session, 2026, 11, 5), context)).toBe(0);

            expect(calendarEventRepo.delete).not.toHaveBeenCalled();
            expect(calendarEventRepo.update).toHaveBeenCalledTimes(1);
            const [delta, existing] = calendarEventRepo.update.mock.calls[0];
            expect(delta.uid).toBe("series");
            expect(delta.version).toBe(4);
            expect(delta.attendees).toBeUndefined(); // the series' own response is untouched
            expect(delta.recurrenceRule.exceptions).toEqual([new Date("2026-10-08T16:00:00.000Z"), new Date("2026-11-05T17:00:00.000Z")]);
            expect(delta.recurrenceRule.byDay).toEqual(["TH"]);
            expect(existing).toBe(series);

            const ics = send.mock.calls[0][0].raw.toString();
            expect(ics).toContain("RECURRENCE-ID:20261105T170000Z");
            expect(ics).not.toContain("RRULE");
            expect(ics).toContain("Subject: Declined: Weekly");
        });

        it("Declining an occurrence that is already an exception changes nothing, and accepting one never touches the series.", async () => {
            const series = makeSeries();
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([series]), update: vi.fn(), delete: vi.fn() };
            const send = vi.fn().mockResolvedValue({ accepted: ["boss@example.com"], rejected: [] });
            const context = makeContext({ calendarEventRepo: calendarEventRepo as any, mailboxRepo: mailboxRepo() as any, mailTransport: { send } as any });

            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", instanceProperty(context.session, 2026, 10, 8), context)).toBe(0);
            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Tent", instanceProperty(context.session, 2026, 10, 15), context)).toBe(0);

            expect(calendarEventRepo.update).not.toHaveBeenCalled();
            expect(calendarEventRepo.delete).not.toHaveBeenCalled();
            expect(send.mock.calls[1][0].raw.toString()).toContain("RECURRENCE-ID:20261015T160000Z");
        });

        it("Handles a series stored with no exceptions list and no usable time zone (UTC).", async () => {
            const series = makeSeries({ timezone: "Not/AZone", recurrenceRule: { freq: "daily", interval: 1 } });
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([series]), update: vi.fn().mockResolvedValue(undefined) };
            const context = makeContext({
                calendarEventRepo: calendarEventRepo as any,
                mailboxRepo: mailboxRepo() as any,
                mailTransport: { send: vi.fn().mockResolvedValue({ accepted: ["boss@example.com"], rejected: [] }) } as any,
            });

            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", instanceProperty(context.session, 2026, 10, 3), context)).toBe(0);

            expect(calendarEventRepo.update.mock.calls[0][0].recurrenceRule.exceptions).toEqual([new Date("2026-10-03T16:00:00.000Z")]);
        });

        it("Matches an exception copy by its date in the event's time zone, not in UTC.", async () => {
            // 20:00 Pacific on 2026-10-08 is already 2026-10-09 in UTC.
            const exception = makeSeries({ uid: "exception", recurrenceRule: undefined, recurrenceId: new Date("2026-10-09T03:00:00.000Z") });
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([makeSeries(), exception]), update: vi.fn().mockResolvedValue(undefined) };
            const context = makeContext({
                calendarEventRepo: calendarEventRepo as any,
                mailboxRepo: mailboxRepo() as any,
                mailTransport: { send: vi.fn().mockResolvedValue({ accepted: ["boss@example.com"], rejected: [] }) } as any,
            });

            await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", instanceProperty(context.session, 2026, 10, 8), context);

            expect(calendarEventRepo.update.mock.calls[0][0].uid).toBe("exception");
        });
    });

    it("Matches a native Outlook GlobalObjectId by the uppercase hex of the whole id (instance date zeroed), in either case.", async () => {
        const blob = Buffer.concat([
            encodeGlobalObjectId("x", new Date("2026-01-01T00:00:00.000Z")).subarray(0, 36),
            Buffer.from([4, 0, 0, 0]),
            Buffer.from([0xde, 0xad, 0xbe, 0xef]),
        ]);
        blob[16] = 0x07;
        blob[17] = 0xea;
        blob[18] = 10;
        blob[19] = 8;
        const zeroed = Buffer.from(blob);
        zeroed.fill(0, 16, 20);
        const hexUid = zeroed.toString("hex").toUpperCase();
        const event = { uid: "evt1", version: 1, icalUid: hexUid.toLowerCase(), attendees: [{ address: "caller@example.com" }] };
        const find = vi.fn().mockImplementation((query: any) => Promise.resolve(query.icalUid === `eq(${hexUid.toLowerCase()})` ? [event] : []));
        const context = makeContext({
            calendarEventRepo: { find, update: vi.fn().mockResolvedValue(undefined) } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const id = assignOrGetNamedPropertyId(context.session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });

        expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", { [String(id)]: blob.toString("base64") }, context)).toBe(0);

        expect(find.mock.calls.map((call) => call[0].icalUid)).toEqual([`eq(${hexUid})`, `eq(${hexUid.toLowerCase()})`]);
        expect((context.calendarEventRepo as any).update).toHaveBeenCalledTimes(1);
    });

    it("Matches the caller via an alias address, not just primarySmtpAddress.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
            icalUid: "evt-uid@example.com",
            attendees: [{ address: "alias@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION }],
        };
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([event]), update: vi.fn().mockResolvedValue(undefined) };
        const context = makeContext({
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "primary@example.com", aliasAddresses: ["alias@example.com"] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect(calendarEventRepo.update).toHaveBeenCalledTimes(1);
    });
    describe("Round 5: lookups and versioned updates", () => {
        const mailboxRepo = () => ({ findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) });

        it("Looks the UID up literally and ignores rows whose icalUid isn't exactly it, so a UID like ne(x) can't pick another meeting.", async () => {
            const other = { uid: "other", version: 1, icalUid: "someone-elses-meeting", attendees: [{ address: "caller@example.com" }] };
            const find = vi.fn().mockResolvedValue([other]);
            const context = makeContext({ calendarEventRepo: { find, update: vi.fn(), delete: vi.fn() } as any, mailboxRepo: mailboxRepo() as any });

            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", globalObjectIdProperty(context.session, "ne(x)"), context)).toBe(0x8004010f);

            expect(find.mock.calls[0][0].icalUid).toBe("eq(ne(x))");
            expect((context.calendarEventRepo as any).delete).not.toHaveBeenCalled();
        });

        it("Bounds a UID over 255 characters the way restapi stores it, and treats a query the repo rejects as no match.", async () => {
            const longUid = "u".repeat(300);
            const { createHash } = await import("crypto");
            const stored = `sha256:${createHash("sha256").update(longUid, "utf8").digest("hex")}`;
            const event = { uid: "evt1", version: 1, icalUid: stored, attendees: [{ address: "caller@example.com" }] };
            const find = vi.fn().mockResolvedValue([event]);
            const context = makeContext({ calendarEventRepo: { find, update: vi.fn().mockResolvedValue(undefined) } as any, mailboxRepo: mailboxRepo() as any });

            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(context.session, longUid), context)).toBe(0);
            expect(find.mock.calls[0][0].icalUid).toBe(`eq(${stored})`);

            const rejecting = makeContext({ calendarEventRepo: { find: vi.fn().mockRejectedValue(new Error("unknown operator")) } as any, mailboxRepo: mailboxRepo() as any });
            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(rejecting.session, "Me"), rejecting)).toBe(0x8004010f);
        });

        it("Updates through an instance of the repo's model class, so the version lock applies to plain Mongo documents.", async () => {
            class Model {
                public constructor(other: object) {
                    Object.assign(this, other);
                }
            }
            const event = { uid: "evt1", version: 3, icalUid: "evt-uid@example.com", attendees: [{ address: "caller@example.com" }] };
            const update = vi.fn().mockResolvedValue(undefined);
            const context = makeContext({ calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update, modelClass: Model } as any, mailboxRepo: mailboxRepo() as any });

            await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", globalObjectIdProperty(context.session, "evt-uid@example.com"), context);

            const existing = update.mock.calls[0][1];
            expect(existing).toBeInstanceOf(Model);
            expect(existing.version).toBe(3);
        });

        it("Declining an occurrence through its exception copy deletes the copy and adds the occurrence to the series' exceptions.", async () => {
            const recurrenceId = new Date("2026-10-08T10:00:00.000Z");
            const series = { uid: "series", version: 2, icalUid: "evt-uid@example.com", recurrenceRule: { freq: "weekly", interval: 1 }, attendees: [{ address: "caller@example.com" }] };
            const exception = { uid: "exception", version: 1, icalUid: "evt-uid@example.com", recurrenceId, attendees: [{ address: "caller@example.com" }] };
            const instanceProperty = (session: MapiSessionContext): Record<string, string> => {
                const id = assignOrGetNamedPropertyId(session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
                const bytes = encodeGlobalObjectId("evt-uid@example.com", new Date());
                bytes[16] = 2026 >> 8;
                bytes[17] = 2026 & 0xff;
                bytes[18] = 10;
                bytes[19] = 8;
                return { [String(id)]: bytes.toString("base64") };
            };
            const repo = { find: vi.fn().mockResolvedValue([series, exception]), update: vi.fn().mockResolvedValue(undefined), delete: vi.fn().mockResolvedValue(undefined) };
            const context = makeContext({ calendarEventRepo: repo as any, mailboxRepo: mailboxRepo() as any });

            expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", instanceProperty(context.session), context)).toBe(0);

            expect(repo.delete).toHaveBeenCalledWith("exception", { ignoreACL: true });
            expect(repo.update).toHaveBeenCalledTimes(1);
            expect(repo.update.mock.calls[0][0]).toEqual({ uid: "series", version: 2, recurrenceRule: { freq: "weekly", interval: 1, exceptions: [recurrenceId] } });

            // Already excluded, or a series that doesn't recur (or isn't in the caller's mailbox): only the copy goes.
            for (const copies of [
                [{ ...series, recurrenceRule: { freq: "weekly", interval: 1, exceptions: [recurrenceId.toISOString()] } }, exception],
                [{ ...series, recurrenceRule: undefined }, exception],
                [exception],
            ]) {
                const again = { find: vi.fn().mockResolvedValue(copies), update: vi.fn(), delete: vi.fn().mockResolvedValue(undefined) };
                const next = makeContext({ calendarEventRepo: again as any, mailboxRepo: mailboxRepo() as any });
                expect(await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", instanceProperty(next.session), next)).toBe(0);
                expect(again.delete).toHaveBeenCalledWith("exception", { ignoreACL: true });
                expect(again.update).not.toHaveBeenCalled();
            }
        });
    });
});
