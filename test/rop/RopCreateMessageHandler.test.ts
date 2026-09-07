///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopCreateMessageHandler } from "../../src/rop/RopCreateMessageHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest({
    logonId = 0,
    inputHandleIndex = 0,
    outputHandleIndex = 1,
    codePageId = 0,
    folderId = 1n,
    associatedFlag = 0,
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt16LE(codePageId);
    writer.writeBigUInt64LE(folderId);
    writer.writeUInt8(associatedFlag);
    return writer.toBuffer();
}

function makeContext(folderIds: Record<string, string>): RopContext {
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    session.folderIds = folderIds;
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session,
        folderRepo: {} as any,
        messageRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopCreateMessageHandler Tests", () => {
    it("Has RopId 0x06.", () => {
        expect(new RopCreateMessageHandler().ropId).toBe(0x06);
    });

    it("Creates a draft message handle against a known FID, returning a well-formed success response.", () => {
        const context = makeContext({ "1": "folder:inbox-uid" });
        const handler = new RopCreateMessageHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, folderId: 1n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x06); // RopId
        expect(response.readUInt8()).toBe(5); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt8()).toBe(0); // HasMessageId
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[5]).toEqual({
            type: "message",
            entityUid: "",
            draftFolderUid: "folder:inbox-uid",
            draftProperties: {},
        });
    });

    it("Returns MAPI_E_NOT_FOUND for an unrecognized FID, without creating a handle.", () => {
        const context = makeContext({});
        const handler = new RopCreateMessageHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({ outputHandleIndex: 5, folderId: 999n })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x06);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0x8004010f);
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[5]).toBeUndefined();
    });
});
