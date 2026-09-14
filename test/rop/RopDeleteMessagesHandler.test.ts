///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopDeleteMessagesHandler } from "../../src/rop/RopDeleteMessagesHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { AuditAction } from "@rapidmx/restapi";

function buildRequest({
    logonId = 0,
    inputHandleIndex = 5,
    wantAsynchronous = 0,
    notifyNonRead = 0,
    messageIds = [] as bigint[],
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(wantAsynchronous);
    writer.writeUInt8(notifyNonRead);
    writer.writeUInt16LE(messageIds.length);
    for (const id of messageIds) {
        writer.writeBigUInt64LE(id);
    }
    return writer.toBuffer();
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: { delete: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: { delete: vi.fn().mockResolvedValue(undefined) } as any,
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

describe("RopDeleteMessagesHandler Tests", () => {
    it("Has RopId 0x1E.", () => {
        expect(new RopDeleteMessagesHandler().ropId).toBe(0x1e);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const context = makeContext();
        const handler = new RopDeleteMessagesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Deletes a message: target and reports PartialCompletion=false.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        context.session.messageIds = { "1": "message:m1" };
        const handler = new RopDeleteMessagesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ messageIds: [1n] })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x1e);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(0); // PartialCompletion
        expect(response.hasMore()).toBe(false);

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m1", { ignoreACL: true });
        expect((context.calendarEventRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Deletes a calendarEvent: target via calendarEventRepo instead of messageRepo.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:cal1" };
        context.session.messageIds = { "1": "calendarEvent:evt1" };
        const handler = new RopDeleteMessagesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ messageIds: [1n] })), writer, context);

        expect((context.calendarEventRepo as any).delete).toHaveBeenCalledWith("evt1", { ignoreACL: true });
        expect((context.messageRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Reports PartialCompletion=true for an unrecognized MessageId, while still deleting the others.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        context.session.messageIds = { "1": "message:m1" };
        const handler = new RopDeleteMessagesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ messageIds: [1n, 999n] })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(1); // PartialCompletion
        expect((context.messageRepo as any).delete).toHaveBeenCalledTimes(1);
    });

    it("Audits a deleted message as MESSAGE_DELETE when the context records audits.", async () => {
        const audit = vi.fn().mockResolvedValue(undefined);
        const context = makeContext({
            audit,
            messageRepo: {
                findOne: vi.fn().mockResolvedValue({ uid: "m1", mailboxUid: "mailbox-1", subject: "Bye", folderUid: "f1" }),
                delete: vi.fn().mockResolvedValue(undefined),
            } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        context.session.messageIds = { "1": "message:m1", "2": "message:gone" };
        (context.messageRepo as any).findOne.mockResolvedValueOnce({ uid: "m1", mailboxUid: "mailbox-1", subject: "Bye", folderUid: "f1" }).mockResolvedValueOnce(undefined);

        await new RopDeleteMessagesHandler().handle(new BufferReader(buildRequest({ messageIds: [1n, 2n] })), new BufferWriter(), context);

        expect((context.messageRepo as any).delete).toHaveBeenCalledTimes(2);
        expect(audit).toHaveBeenCalledTimes(1); // a message that no longer exists has nothing to audit
        expect(audit).toHaveBeenCalledWith({
            action: AuditAction.MESSAGE_DELETE,
            targetType: "Message",
            targetUid: "m1",
            mailboxUid: "mailbox-1",
            details: { subject: "Bye", folderUid: "f1" },
        });
    });

    it("Handles an empty MessageIds list, reporting success with PartialCompletion=false.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const handler = new RopDeleteMessagesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(0);
        expect((context.messageRepo as any).delete).not.toHaveBeenCalled();
    });
});
