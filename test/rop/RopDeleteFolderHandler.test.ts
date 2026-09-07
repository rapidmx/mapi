///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopDeleteFolderHandler } from "../../src/rop/RopDeleteFolderHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const DEL_MESSAGES = 0x01;
const DEL_FOLDERS = 0x04;
const DELETE_HARD_DELETE = 0x10;

function buildRequest({ logonId = 0, inputHandleIndex = 5, deleteFolderFlags = 0, folderId = 1n }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(deleteFolderFlags);
    writer.writeBigUInt64LE(folderId);
    return writer.toBuffer();
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        messageRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        calendarEventRepo: { find: vi.fn().mockResolvedValue([]), delete: vi.fn().mockResolvedValue(undefined) } as any,
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

describe("RopDeleteFolderHandler Tests", () => {
    it("Has RopId 0x1D.", () => {
        expect(new RopDeleteFolderHandler().ropId).toBe(0x1d);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const context = makeContext();
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Returns MAPI_E_NOT_FOUND for an unrecognized FolderId.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ folderId: 999n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x8004010f);
    });

    it("Deletes an empty folder with no flags set.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x1d);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt8()).toBe(0); // PartialCompletion
        expect(response.hasMore()).toBe(false);

        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true, purge: false });
    });

    it("Returns MAPI_E_INVALID_OBJECT for a non-empty folder (has messages) when DEL_MESSAGES isn't set, without deleting anything.", async () => {
        const context = makeContext({
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
        expect((context.messageRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Returns MAPI_E_INVALID_OBJECT for a folder with calendar events when DEL_MESSAGES isn't set.", async () => {
        const context = makeContext({
            calendarEventRepo: { find: vi.fn().mockResolvedValue([{ uid: "evt1" }]), delete: vi.fn() } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Deletes messages and calendar events, then the folder, when DEL_MESSAGES is set.", async () => {
        const context = makeContext({
            messageRepo: { find: vi.fn().mockResolvedValue([{ uid: "m1" }, { uid: "m2" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
            calendarEventRepo: { find: vi.fn().mockResolvedValue([{ uid: "evt1" }]), delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_MESSAGES })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m1", { ignoreACL: true, purge: false });
        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("m2", { ignoreACL: true, purge: false });
        expect((context.calendarEventRepo as any).delete).toHaveBeenCalledWith("evt1", { ignoreACL: true, purge: false });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true, purge: false });
    });

    it("Returns MAPI_E_INVALID_OBJECT for a folder with subfolders when DEL_FOLDERS isn't set.", async () => {
        const context = makeContext({
            folderRepo: {
                find: vi.fn().mockResolvedValue([{ uid: "child1" }]),
                delete: vi.fn(),
            } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect((context.folderRepo as any).delete).not.toHaveBeenCalled();
    });

    it("Recursively deletes subfolders (and their own messages/events) when DEL_FOLDERS is set.", async () => {
        const folderFind = vi.fn().mockImplementation(({ parentFolderUid }: { parentFolderUid: string }) => {
            if (parentFolderUid === "target") return Promise.resolve([{ uid: "child1" }]);
            if (parentFolderUid === "child1") return Promise.resolve([{ uid: "grandchild1" }]);
            return Promise.resolve([]);
        });
        const messageFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "child1") return Promise.resolve([{ uid: "child-msg" }]);
            return Promise.resolve([]);
        });
        const eventFind = vi.fn().mockImplementation(({ folderUid }: { folderUid: string }) => {
            if (folderUid === "grandchild1") return Promise.resolve([{ uid: "grandchild-evt" }]);
            return Promise.resolve([]);
        });
        const context = makeContext({
            folderRepo: { find: folderFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            messageRepo: { find: messageFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
            calendarEventRepo: { find: eventFind, delete: vi.fn().mockResolvedValue(undefined) } as any,
        });
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DEL_FOLDERS })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);

        expect((context.messageRepo as any).delete).toHaveBeenCalledWith("child-msg", { ignoreACL: true, purge: false });
        expect((context.calendarEventRepo as any).delete).toHaveBeenCalledWith("grandchild-evt", { ignoreACL: true, purge: false });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("grandchild1", { ignoreACL: true, purge: false });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("child1", { ignoreACL: true, purge: false });
        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true, purge: false });
    });

    it("Passes purge:true to every delete call when DELETE_HARD_DELETE is set.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:parent" };
        context.session.folderIds = { "1": "folder:target" };
        const handler = new RopDeleteFolderHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ deleteFolderFlags: DELETE_HARD_DELETE })), writer, context);

        expect((context.folderRepo as any).delete).toHaveBeenCalledWith("target", { ignoreACL: true, purge: true });
    });
});
