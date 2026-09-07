///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { simpleParser } from "mailparser";
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopSubmitMessageHandler } from "../../src/rop/RopSubmitMessageHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { AvVerdict, FolderType, SpamVerdict } from "@rapidmx/restapi";
import { InMemoryBlobStore } from "../testDoubles.js";

function buildRequest({ logonId = 0, inputHandleIndex = 5, submitFlags = 0 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(submitFlags);
    return writer.toBuffer();
}

function makeScanPipeline() {
    return {
        run: vi.fn().mockResolvedValue({
            spam: { verdict: SpamVerdict.CLEAN },
            av: { verdict: AvVerdict.CLEAN },
            attachments: [],
        }),
    };
}

function makeMailTransport() {
    return { send: vi.fn().mockResolvedValue({ accepted: ["placeholder"], rejected: [], messageId: "x" }) };
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { find: vi.fn().mockResolvedValue([{ uid: "sent-uid", type: FolderType.SENT_ITEMS }]) } as any,
        messageRepo: { create: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: {} as any,
        mailboxRepo: { findOne: vi.fn().mockResolvedValue({ uid: "mailbox-1", primarySmtpAddress: "owner@example.com" }) } as any,
        folderClass: class TestFolder {
            public constructor(data: any) {
                Object.assign(this, data);
            }
        },
        messageClass: class TestMessage {
            public constructor(data: any) {
                Object.assign(this, data);
            }
        },
        calendarEventClass: {} as any,
        scanPipeline: makeScanPipeline() as any,
        mailTransport: makeMailTransport() as any,
        blobStore: new InMemoryBlobStore(),
        ...overrides,
    };
}

describe("RopSubmitMessageHandler Tests", () => {
    it("Has RopId 0x32.", () => {
        expect(new RopSubmitMessageHandler().ropId).toBe(0x32);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a message handle.", async () => {
        const context = makeContext();
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns MAPI_E_INVALID_OBJECT when the draft has no recipients at all.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "55": "No Recipients" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(context.mailTransport.send).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT when the draft handle has no draftProperties at all.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1" };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Succeeds with an empty Subject and body when neither was ever set, only a recipient.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "3588": "to@example.com" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.subject).toBeUndefined();
        expect((parsed.text ?? "").trim()).toBe("");
    });

    it("Returns MAPI_E_INVALID_OBJECT when the mailbox has no resolvable primarySmtpAddress.", async () => {
        const context = makeContext({ mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any });
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "3588": "to@example.com" },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Builds and sends a real MIME message from Subject/DisplayTo/DisplayCc/inline Body, then saves a Sent Items copy.", async () => {
        const context = makeContext();
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: {
                "55": "Hello From MAPI",
                "3588": "to@example.com",
                "3587": "cc@example.com",
                "3586": "bcc@example.com",
                "4096": "This is the message body.",
            },
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x32);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success
        expect(response.hasMore()).toBe(false);

        const scanPipeline = context.scanPipeline as any;
        expect(scanPipeline.run).toHaveBeenCalledTimes(1);
        const [rawSent, envelope] = scanPipeline.run.mock.calls[0];
        expect(envelope).toEqual({ from: "owner@example.com", to: ["to@example.com", "cc@example.com", "bcc@example.com"] });

        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.subject).toBe("Hello From MAPI");
        expect(parsed.text?.trim()).toBe("This is the message body.");
        expect(parsed.from?.value[0]?.address).toBe("owner@example.com");

        expect(context.mailTransport.send).toHaveBeenCalledTimes(1);

        const messageRepo = context.messageRepo as any;
        expect(messageRepo.create).toHaveBeenCalledTimes(1);
        const [savedMessage, options] = messageRepo.create.mock.calls[0];
        expect(savedMessage.folderUid).toBe("sent-uid");
        expect(savedMessage.subject).toBe("Hello From MAPI");
        expect(savedMessage.recipients).toEqual([
            { address: "to@example.com", type: "to" },
            { address: "cc@example.com", type: "cc" },
            { address: "bcc@example.com", type: "bcc" },
        ]);
        expect(options).toEqual({ ignoreACL: true });
    });

    it("Prefers a RopWriteStream-accumulated body over an inline PidTagBody draftProperty.", async () => {
        const context = makeContext();
        const streamText = "Body written via RopWriteStream.";
        const streamBytes = Buffer.concat([Buffer.from(streamText, "utf16le"), Buffer.from([0, 0])]);
        context.session.handles[5] = {
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:f1",
            draftProperties: { "55": "Subj", "3588": "to@example.com", "4096": "This inline body must be ignored." },
        };
        context.session.handles[6] = {
            type: "stream",
            entityUid: "",
            writeTargetHandleIndex: 5,
            writeBufferBase64: streamBytes.toString("base64"),
        };
        const handler = new RopSubmitMessageHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const scanPipeline = context.scanPipeline as any;
        const [rawSent] = scanPipeline.run.mock.calls[0];
        const parsed = await simpleParser(rawSent as Buffer);
        expect(parsed.text?.trim()).toBe(streamText);
    });

    describe("Calendar branch (PidTagMessageClass starts with IPM.Appointment)", () => {
        function makeCalendarEvent(overrides: Record<string, unknown> = {}) {
            return {
                uid: "evt1",
                icalUid: "abc-123@mapi",
                sequence: 0,
                title: "Standup",
                location: "Room 1",
                startDate: new Date("2026-09-07T14:00:00.000Z"),
                endDate: new Date("2026-09-07T15:00:00.000Z"),
                attendees: [],
                ...overrides,
            };
        }

        it("Returns MAPI_E_INVALID_OBJECT when the handle's entityUid was never assigned a real calendarEvent (Save never succeeded as an Appointment).", async () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: { "26": "IPM.Appointment" } };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
        });

        it("Returns MAPI_E_INVALID_OBJECT when the referenced CalendarEvent has since vanished.", async () => {
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:gone",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0x80070005);
        });

        it("Succeeds without sending anything when the event has no attendees (a private, non-meeting appointment).", async () => {
            const event = makeCalendarEvent({ attendees: [] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0); // ReturnValue - success
            const scanPipeline = context.scanPipeline as any;
            expect(scanPipeline.run).not.toHaveBeenCalled();
        });

        it("Skips sending when the mailbox has no resolvable envelope-from address, even with attendees present.", async () => {
            const event = makeCalendarEvent({ attendees: [{ address: "attendee@example.com" }] });
            const context = makeContext({
                calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any,
                mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
            });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0); // still success - the appointment itself already exists
            const scanPipeline = context.scanPipeline as any;
            expect(scanPipeline.run).not.toHaveBeenCalled();
        });

        it("Sends a real iCalendar METHOD:REQUEST invite to every attendee, without creating a Sent Items message.", async () => {
            const event = makeCalendarEvent({
                attendees: [{ address: "attendee1@example.com" }, { address: "attendee2@example.com" }],
            });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            expect(response.readUInt8()).toBe(0x32);
            expect(response.readUInt8()).toBe(5);
            expect(response.readUInt32LE()).toBe(0); // ReturnValue - success

            const scanPipeline = context.scanPipeline as any;
            expect(scanPipeline.run).toHaveBeenCalledTimes(1);
            const [rawSent, envelope] = scanPipeline.run.mock.calls[0];
            expect(envelope).toEqual({ from: "owner@example.com", to: ["attendee1@example.com", "attendee2@example.com"] });

            const raw = (rawSent as Buffer).toString("utf-8");
            expect(raw).toContain("text/calendar; charset=utf-8; method=REQUEST");
            expect(raw).toContain("BEGIN:VEVENT");
            expect(raw).toContain("UID:abc-123@mapi");
            expect(raw).toContain("SUMMARY:Standup");
            expect(raw).toContain("LOCATION:Room 1");
            expect(raw).toContain("ORGANIZER:mailto:owner@example.com");
            expect(raw).toContain("ATTENDEE;RSVP=TRUE:mailto:attendee1@example.com");
            expect(raw).toContain("ATTENDEE;RSVP=TRUE:mailto:attendee2@example.com");

            expect(context.mailTransport.send).toHaveBeenCalledTimes(1);
            expect((context.messageRepo as any).create).not.toHaveBeenCalled();
        });

        it("Ignores an unrecognized IPM.Appointment.* subclass exactly the same way (prefix match, not exact match).", async () => {
            const event = makeCalendarEvent({ attendees: [] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment.SomeSubclass" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            expect(response.readUInt32LE()).toBe(0);
        });

        it("Omits LOCATION from the invite when the event has none.", async () => {
            const event = makeCalendarEvent({ location: undefined, attendees: [{ address: "attendee@example.com" }] });
            const context = makeContext({ calendarEventRepo: { findOne: vi.fn().mockResolvedValue(event) } as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "calendarEvent:evt1",
                draftProperties: { "26": "IPM.Appointment" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const scanPipeline = context.scanPipeline as any;
            const [rawSent] = scanPipeline.run.mock.calls[0];
            const raw = (rawSent as Buffer).toString("utf-8");
            expect(raw).not.toContain("LOCATION:");
        });
    });

    describe("Meeting-response branch (PidTagMessageClass starts with IPM.Schedule.Meeting.Resp.)", () => {
        it("Dispatches to submitMeetingResponse and reports success, without touching the mail compose/send path.", async () => {
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([]) };
            const context = makeContext({ calendarEventRepo: calendarEventRepo as any });
            context.session.handles[5] = {
                type: "message",
                entityUid: "",
                draftProperties: { "26": "IPM.Schedule.Meeting.Resp.Pos" },
            };
            const handler = new RopSubmitMessageHandler();
            const writer = new BufferWriter();

            await handler.handle(new BufferReader(buildRequest({})), writer, context);

            const response = new BufferReader(writer.toBuffer());
            expect(response.readUInt8()).toBe(0x32);
            expect(response.readUInt8()).toBe(5);
            expect(response.readUInt32LE()).toBe(0); // ReturnValue - success
            expect(response.hasMore()).toBe(false);

            // No PidLidGlobalObjectId was ever set on this draft, so submitMeetingResponse should have no-op'd
            // before even querying the calendar event repo - confirming real dispatch happened either way.
            expect(context.mailTransport.send).not.toHaveBeenCalled();
            expect((context.messageRepo as any).create).not.toHaveBeenCalled();
        });
    });
});
