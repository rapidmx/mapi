///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopGetHierarchyTableHandler } from "../../src/rop/RopGetHierarchyTableHandler.js";
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

describe("RopGetHierarchyTableHandler Tests", () => {
    it("Has RopId 0x04.", () => {
        expect(new RopGetHierarchyTableHandler().ropId).toBe(0x04);
    });

    it("Creates a table handle listing the folder's children, resolved via the folder repo.", async () => {
        const folderRepo = {
            find: vi.fn().mockResolvedValue([
                { uid: "child1", parentFolderUid: "top1", mailboxUid: "mailbox-1" },
                { uid: "child2", parentFolderUid: "top1", mailboxUid: "mailbox-1" },
                { uid: "other", parentFolderUid: "somewhere-else", mailboxUid: "mailbox-1" },
            ]),
        };
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        session.handles[5] = { type: "folder", entityUid: "folder:top1" };
        const context: RopContext = { mailboxUid: "mailbox-1", userUid: "user-1", session, folderRepo: folderRepo as any, messageRepo: {} as any, mailboxRepo: {} as any, folderClass: {} as any, messageClass: {} as any, scanPipeline: {} as any, mailTransport: {} as any, blobStore: {} as any };

        const handler = new RopGetHierarchyTableHandler();
        const writer = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x04);
        expect(response.readUInt8()).toBe(6); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.hasMore()).toBe(false);

        expect(session.handles[6]).toEqual({
            type: "table",
            entityUid: "folder:top1",
            rows: ["folder:child1", "folder:child2"],
            cursor: 0,
            generation: expect.any(String),
        });
        // The mailbox's folders are fetched with explicit paging and sort, not the repo's 100-row default.
        expect(folderRepo.find).toHaveBeenCalledWith(
            { mailboxUid: "mailbox-1", sort: { name: "ASC", uid: "ASC" }, limit: 1000, page: 0 },
            { ignoreACL: true, limit: 1000, page: 0 },
        );
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder handle.", async () => {
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        const context: RopContext = { mailboxUid: "mailbox-1", userUid: "user-1", session, folderRepo: {} as any, messageRepo: {} as any, mailboxRepo: {} as any, folderClass: {} as any, messageClass: {} as any, scanPipeline: {} as any, mailTransport: {} as any, blobStore: {} as any };
        const handler = new RopGetHierarchyTableHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ inputHandleIndex: 99 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(session.handles[6]).toBeUndefined();
    });
});
