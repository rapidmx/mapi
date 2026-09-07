///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, writePropertyTag } from "../../src/codec/PropertyValue.js";
import { RopFastTransferSourceCopyToHandler } from "../../src/rop/RopFastTransferSourceCopyToHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { FolderType } from "@rapidmx/restapi";

function buildRequest({
    logonId = 0,
    inputHandleIndex = 5,
    outputHandleIndex = 6,
    level = 0,
    copyFlags = 0,
    sendOptions = 0,
    excludedTags = [] as { propertyId: number; propertyType: PropertyType }[],
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt8(level);
    writer.writeUInt32LE(copyFlags);
    writer.writeUInt8(sendOptions);
    writer.writeUInt16LE(excludedTags.length);
    for (const tag of excludedTags) {
        writePropertyTag(writer, tag);
    }
    return writer.toBuffer();
}

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "f1", type: FolderType.INBOX }), find: vi.fn().mockResolvedValue([]) } as any,
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

describe("RopFastTransferSourceCopyToHandler Tests", () => {
    it("Has RopId 0x4D.", () => {
        expect(new RopFastTransferSourceCopyToHandler().ropId).toBe(0x4d);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder or message handle.", async () => {
        const context = makeContext();
        const handler = new RopFastTransferSourceCopyToHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Creates a fastTransfer output handle carrying the built stream, for a folder handle.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const handler = new RopFastTransferSourceCopyToHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x4d);
        expect(response.readUInt8()).toBe(6);
        expect(response.readUInt32LE()).toBe(0);
        expect(response.hasMore()).toBe(false);

        const outputHandle = context.session.handles[6];
        expect(outputHandle?.type).toBe("fastTransfer");
        expect(outputHandle?.entityUid).toBe("folder:f1");
        expect(outputHandle?.transferPosition).toBe(0);
        expect(Buffer.from(outputHandle?.transferBufferBase64 ?? "", "base64").length).toBeGreaterThan(0);
    });

    it("Creates a fastTransfer output handle for a message handle.", async () => {
        const context = makeContext({ messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }) } as any });
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const handler = new RopFastTransferSourceCopyToHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        expect(context.session.handles[6]?.type).toBe("fastTransfer");
    });

    it("Excludes the requested PropertyTags from the default column set.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const handler = new RopFastTransferSourceCopyToHandler();
        const writer = new BufferWriter();

        await handler.handle(
            new BufferReader(buildRequest({ excludedTags: [{ propertyId: 0x3001, propertyType: PropertyType.PtypString }] })),
            writer,
            context,
        );

        const outputHandle = context.session.handles[6];
        const buffer = Buffer.from(outputHandle?.transferBufferBase64 ?? "", "base64");
        // DisplayName (0x3001) excluded - the stream should be shorter than the default (all 3 folder columns).
        const contextWithoutExclusion = makeContext();
        contextWithoutExclusion.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const writer2 = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({})), writer2, contextWithoutExclusion);
        const fullBuffer = Buffer.from(contextWithoutExclusion.session.handles[6]?.transferBufferBase64 ?? "", "base64");
        expect(buffer.length).toBeLessThan(fullBuffer.length);
    });
});
