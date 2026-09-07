///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader } from "../../src/codec/BufferCursor.js";
import { PropertyType, readTaggedPropertyValue } from "../../src/codec/PropertyValue.js";
import { buildFastTransferStream } from "../../src/rop/FastTransferStream.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { FolderType } from "@rapidmx/restapi";

const MARKER_START_TOP_FLD = 0x40090003;
const MARKER_END_FOLDER = 0x400b0003;
const MARKER_START_MESSAGE = 0x400c0003;
const MARKER_END_MESSAGE = 0x400d0003;

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { findOne: vi.fn().mockResolvedValue(undefined), find: vi.fn().mockResolvedValue([]) } as any,
        messageRepo: { find: vi.fn().mockResolvedValue([]), findOne: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: { find: vi.fn().mockResolvedValue([]), findOne: vi.fn().mockResolvedValue(undefined) } as any,
        mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
        ...overrides,
    };
}

describe("FastTransferStream Tests", () => {
    describe("Folder targets", () => {
        it("Wraps a folder's default columns and its (non-calendar) messages in StartTopFld/EndFolder + StartMessage/EndMessage markers.", async () => {
            const context = makeContext({
                folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox", type: FolderType.INBOX }), find: vi.fn().mockResolvedValue([]) } as any,
                messageRepo: {
                    find: vi.fn().mockResolvedValue([{ uid: "m1", folderUid: "f1" }]),
                    findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hello", flags: { read: true } }),
                } as any,
            });
            const handle = { type: "folder" as const, entityUid: "folder:f1" };

            const buffer = await buildFastTransferStream(handle, context);
            const reader = new BufferReader(buffer);

            expect(reader.readUInt32LE()).toBe(MARKER_START_TOP_FLD);
            const displayName = readTaggedPropertyValue(reader);
            expect(displayName).toEqual({ propertyId: 0x3001, propertyType: PropertyType.PtypString, value: "Inbox" });
            readTaggedPropertyValue(reader); // ContentCount
            readTaggedPropertyValue(reader); // ContentUnreadCount

            expect(reader.readUInt32LE()).toBe(MARKER_START_MESSAGE);
            const subject = readTaggedPropertyValue(reader);
            expect(subject).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hello" });
            readTaggedPropertyValue(reader); // MessageFlags
            readTaggedPropertyValue(reader); // HasAttachments
            readTaggedPropertyValue(reader); // MessageDeliveryTime
            expect(reader.readUInt32LE()).toBe(MARKER_END_MESSAGE);

            expect(reader.readUInt32LE()).toBe(MARKER_END_FOLDER);
            expect(reader.hasMore()).toBe(false);
        });

        it("Uses calendarEvent rows (Subject only) instead of messages for a CALENDAR-type folder.", async () => {
            const context = makeContext({
                folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "cal1", type: FolderType.CALENDAR }), find: vi.fn().mockResolvedValue([]) } as any,
                calendarEventRepo: {
                    find: vi.fn().mockResolvedValue([{ uid: "evt1", folderUid: "cal1" }]),
                    findOne: vi.fn().mockResolvedValue({ uid: "evt1", title: "Standup" }),
                } as any,
                messageRepo: { find: vi.fn(), findOne: vi.fn() } as any,
            });
            const handle = { type: "folder" as const, entityUid: "folder:cal1" };

            const buffer = await buildFastTransferStream(handle, context);
            const reader = new BufferReader(buffer);

            reader.readUInt32LE(); // StartTopFld
            readTaggedPropertyValue(reader); // DisplayName
            readTaggedPropertyValue(reader); // ContentCount
            readTaggedPropertyValue(reader); // ContentUnreadCount
            expect(reader.readUInt32LE()).toBe(MARKER_START_MESSAGE);
            expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Standup" });
            expect(reader.readUInt32LE()).toBe(MARKER_END_MESSAGE);
            expect(reader.readUInt32LE()).toBe(MARKER_END_FOLDER);
            expect(reader.hasMore()).toBe(false);
            expect((context.messageRepo as any).find).not.toHaveBeenCalled();
        });

        it("Emits an empty message list (no rows) for a virtual folder target (not '\"folder:\"'-prefixed).", async () => {
            const context = makeContext();
            const handle = { type: "folder" as const, entityUid: "virtual:root" };

            const buffer = await buildFastTransferStream(handle, context);
            const reader = new BufferReader(buffer);

            expect(reader.readUInt32LE()).toBe(MARKER_START_TOP_FLD);
            readTaggedPropertyValue(reader);
            readTaggedPropertyValue(reader);
            readTaggedPropertyValue(reader);
            expect(reader.readUInt32LE()).toBe(MARKER_END_FOLDER);
            expect(reader.hasMore()).toBe(false);
        });

        it("Applies excludePropertyIds against the default folder/message columns.", async () => {
            const context = makeContext({
                folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "f1", type: FolderType.INBOX }), find: vi.fn().mockResolvedValue([]) } as any,
                messageRepo: {
                    find: vi.fn().mockResolvedValue([{ uid: "m1" }]),
                    findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }),
                } as any,
            });
            const handle = { type: "folder" as const, entityUid: "folder:f1" };

            const buffer = await buildFastTransferStream(handle, context, {
                excludePropertyIds: new Set([0x3001, 0x0037]), // DisplayName, Subject
            });
            const reader = new BufferReader(buffer);

            reader.readUInt32LE(); // StartTopFld
            const first = readTaggedPropertyValue(reader);
            expect(first.propertyId).toBe(0x3602); // ContentCount, DisplayName excluded
            readTaggedPropertyValue(reader); // ContentUnreadCount
            reader.readUInt32LE(); // StartMessage
            const messageFirst = readTaggedPropertyValue(reader);
            expect(messageFirst.propertyId).toBe(0x0e07); // MessageFlags, Subject excluded
        });

        it("Uses an explicit include column list uniformly at both folder- and message-level, ignoring the default sets entirely.", async () => {
            const context = makeContext({
                folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "f1", type: FolderType.INBOX }), find: vi.fn().mockResolvedValue([]) } as any,
                messageRepo: {
                    find: vi.fn().mockResolvedValue([{ uid: "m1" }]),
                    findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }),
                } as any,
            });
            const handle = { type: "folder" as const, entityUid: "folder:f1" };
            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }];

            const buffer = await buildFastTransferStream(handle, context, { columns });
            const reader = new BufferReader(buffer);

            reader.readUInt32LE(); // StartTopFld
            expect(readTaggedPropertyValue(reader).propertyId).toBe(0x0037); // Subject used even at folder level
            reader.readUInt32LE(); // StartMessage
            expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hi" });
            reader.readUInt32LE(); // EndMessage
            reader.readUInt32LE(); // EndFolder
            expect(reader.hasMore()).toBe(false);
        });
    });

    describe("Message/calendarEvent targets", () => {
        it("Emits a bare propList (no markers) for a message: target.", async () => {
            const context = makeContext({
                messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi", flags: { read: false } }) } as any,
            });
            const handle = { type: "message" as const, entityUid: "message:m1" };

            const buffer = await buildFastTransferStream(handle, context);
            const reader = new BufferReader(buffer);

            expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hi" });
            readTaggedPropertyValue(reader); // MessageFlags
            readTaggedPropertyValue(reader); // HasAttachments
            readTaggedPropertyValue(reader); // MessageDeliveryTime
            expect(reader.hasMore()).toBe(false);
        });

        it("Emits only Subject (no named properties) for a calendarEvent: target.", async () => {
            const context = makeContext({
                calendarEventRepo: { findOne: vi.fn().mockResolvedValue({ uid: "evt1", title: "Standup" }) } as any,
            });
            const handle = { type: "message" as const, entityUid: "calendarEvent:evt1" };

            const buffer = await buildFastTransferStream(handle, context);
            const reader = new BufferReader(buffer);

            expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Standup" });
            expect(reader.hasMore()).toBe(false);
        });

        it("Uses an explicit include column list for a message: target.", async () => {
            const context = makeContext({
                messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", flags: { read: true } }) } as any,
            });
            const handle = { type: "message" as const, entityUid: "message:m1" };
            const columns = [{ propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 }];

            const buffer = await buildFastTransferStream(handle, context, { columns });
            const reader = new BufferReader(buffer);

            expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32, value: 1 });
            expect(reader.hasMore()).toBe(false);
        });
    });
});
