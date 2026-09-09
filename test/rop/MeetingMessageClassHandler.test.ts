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
    it("Does nothing for an unrecognized message-class suffix.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn() } as any });
        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Bogus", {}, context);
        expect((context.calendarEventRepo as any).find).not.toHaveBeenCalled();
    });

    it("Does nothing when the draft never set PidLidGlobalObjectId.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn() } as any });
        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", {}, context);
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

        expect((context.calendarEventRepo as any).find).toHaveBeenCalledWith({ icalUid: "evt-uid@example.com" }, { ignoreACL: true, limit: 1 });
    });

    it("Does nothing when no CalendarEvent matches the decoded icalUid.", async () => {
        const context = makeContext({ calendarEventRepo: { find: vi.fn().mockResolvedValue([]) } as any });
        const properties = globalObjectIdProperty(context.session, "unknown@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect((context.calendarEventRepo as any).find).toHaveBeenCalledWith({ icalUid: "unknown@example.com" }, { ignoreACL: true, limit: 1 });
    });

    it("Does nothing when the caller's mailbox can't be resolved.", async () => {
        const event = { uid: "evt1", version: 1, attendees: [{ address: "caller@example.com" }] };
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update: vi.fn() } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect((context.calendarEventRepo as any).update).not.toHaveBeenCalled();
    });

    it("Does nothing when the caller isn't actually an attendee of the event.", async () => {
        const event = { uid: "evt1", version: 1, attendees: [{ address: "someone-else@example.com" }] };
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([event]), update: vi.fn() } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Pos", properties, context);

        expect((context.calendarEventRepo as any).update).not.toHaveBeenCalled();
    });

    it("Updates the caller's own Attendee.responseStatus to ACCEPTED for .Resp.Pos, matched by primarySmtpAddress (case-insensitive).", async () => {
        const event = {
            uid: "evt1",
            version: 2,
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

    it("Updates to DECLINED for .Resp.Neg.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
            attendees: [{ address: "caller@example.com", responseStatus: AttendeeResponseStatus.NEEDS_ACTION }],
        };
        const calendarEventRepo = { find: vi.fn().mockResolvedValue([event]), update: vi.fn().mockResolvedValue(undefined) };
        const context = makeContext({
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "caller@example.com", aliasAddresses: [] }) } as any,
        });
        const properties = globalObjectIdProperty(context.session, "evt-uid@example.com");

        await submitMeetingResponse("IPM.Schedule.Meeting.Resp.Neg", properties, context);

        const [delta] = calendarEventRepo.update.mock.calls[0];
        expect(delta.attendees[0].responseStatus).toBe(AttendeeResponseStatus.DECLINED);
    });

    it("Matches the caller via an alias address, not just primarySmtpAddress.", async () => {
        const event = {
            uid: "evt1",
            version: 1,
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
});
