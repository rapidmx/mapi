///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopGetContentsTableHandler } from "../../src/rop/RopGetContentsTableHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({ logonId = 0, inputHandleIndex = 5, outputHandleIndex = 6, tableFlags = 0 }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt8(tableFlags);
    return writer.toBuffer();
}

describe("RopGetContentsTableHandler Tests", () => {
    it("Has RopId 0x05.", () => {
        expect(new RopGetContentsTableHandler().ropId).toBe(0x05);
    });

    it("Creates a table handle listing the folder's messages, resolved via the message repo.", async () => {
        const messageRepo = {
            find: vi.fn().mockResolvedValue([
                { uid: "msg1", folderUid: "top1" },
                { uid: "msg2", folderUid: "top1" },
            ]),
        };
        const folderRepo = { findOne: vi.fn().mockResolvedValue({ uid: "top1", type: "inbox" }) };
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        session.handles[5] = { type: "folder", entityUid: "folder:top1" };
        const context: RopContext = {
            mailboxUid: "mailbox-1",
            userUid: "user-1",
            session,
            folderRepo: folderRepo as any,
            messageRepo: messageRepo as any,
            calendarEventRepo: {} as any,
            mailboxRepo: {} as any,
            folderClass: {} as any,
            messageClass: {} as any,
            calendarEventClass: {} as any,
            scanPipeline: {} as any,
            mailTransport: {} as any,
            blobStore: {} as any,
        };

        const handler = new RopGetContentsTableHandler();
        const writer = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x05);
        expect(response.readUInt8()).toBe(6);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.hasMore()).toBe(false);

        expect(session.handles[6]).toEqual({
            type: "table",
            entityUid: "folder:top1",
            rows: ["message:msg1", "message:msg2"],
            cursor: 0,
        });
        expect(folderRepo.findOne).toHaveBeenCalledWith("top1", { ignoreACL: true });
        expect(messageRepo.find).toHaveBeenCalledWith({ folderUid: "top1" }, { ignoreACL: true });
    });

    it("Creates a table handle listing a CALENDAR folder's events, resolved via the calendar event repo instead of the message repo.", async () => {
        const messageRepo = { find: vi.fn() };
        const calendarEventRepo = {
            find: vi.fn().mockResolvedValue([
                { uid: "evt1", folderUid: "cal1" },
                { uid: "evt2", folderUid: "cal1" },
            ]),
        };
        const folderRepo = { findOne: vi.fn().mockResolvedValue({ uid: "cal1", type: "calendar" }) };
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        session.handles[5] = { type: "folder", entityUid: "folder:cal1" };
        const context: RopContext = {
            mailboxUid: "mailbox-1",
            userUid: "user-1",
            session,
            folderRepo: folderRepo as any,
            messageRepo: messageRepo as any,
            calendarEventRepo: calendarEventRepo as any,
            mailboxRepo: {} as any,
            folderClass: {} as any,
            messageClass: {} as any,
            calendarEventClass: {} as any,
            scanPipeline: {} as any,
            mailTransport: {} as any,
            blobStore: {} as any,
        };

        const handler = new RopGetContentsTableHandler();
        const writer = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect(session.handles[6]?.rows).toEqual(["calendarEvent:evt1", "calendarEvent:evt2"]);
        expect(calendarEventRepo.find).toHaveBeenCalledWith({ folderUid: "cal1" }, { ignoreACL: true });
        expect(messageRepo.find).not.toHaveBeenCalled();
    });

    it("Returns an empty table for a virtual folder, without querying the message repo.", async () => {
        const messageRepo = { find: vi.fn() };
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        session.handles[5] = { type: "folder", entityUid: "virtual:root" };
        const context: RopContext = {
            mailboxUid: "mailbox-1",
            userUid: "user-1",
            session,
            folderRepo: {} as any,
            messageRepo: messageRepo as any,
            calendarEventRepo: {} as any,
            mailboxRepo: {} as any,
            folderClass: {} as any,
            messageClass: {} as any,
            calendarEventClass: {} as any,
            scanPipeline: {} as any,
            mailTransport: {} as any,
            blobStore: {} as any,
        };

        const handler = new RopGetContentsTableHandler();
        const writer = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect(session.handles[6]?.rows).toEqual([]);
        expect(messageRepo.find).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        const context: RopContext = {
            mailboxUid: "mailbox-1",
            userUid: "user-1",
            session,
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
        };
        const handler = new RopGetContentsTableHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ inputHandleIndex: 99 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(session.handles[6]).toBeUndefined();
    });
});
