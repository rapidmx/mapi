///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { MAX_PROPERTY_TAG_COUNT, PropertyType, readTaggedPropertyValue, writePropertyTag } from "../../src/codec/PropertyValue.js";
import { MAX_FAST_TRANSFER_BYTES } from "../../src/rop/FastTransferStream.js";
import { handleDataCache, handleDataKey } from "../../src/rop/HandleDataCache.js";
import { RopFastTransferSourceCopyPropertiesHandler } from "../../src/rop/RopFastTransferSourceCopyPropertiesHandler.js";
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
    includedTags = [] as { propertyId: number; propertyType: PropertyType }[],
}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt8(level);
    writer.writeUInt8(copyFlags);
    writer.writeUInt8(sendOptions);
    writer.writeUInt16LE(includedTags.length);
    for (const tag of includedTags) {
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

describe("RopFastTransferSourceCopyPropertiesHandler Tests", () => {
    it("Has RopId 0x69.", () => {
        expect(new RopFastTransferSourceCopyPropertiesHandler().ropId).toBe(0x69);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a folder or message handle.", async () => {
        const context = makeContext();
        const handler = new RopFastTransferSourceCopyPropertiesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Builds a stream using only the explicitly-requested include-list columns, for a message handle.", async () => {
        const context = makeContext({ messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }) } as any });
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const handler = new RopFastTransferSourceCopyPropertiesHandler();
        const writer = new BufferWriter();

        await handler.handle(
            new BufferReader(buildRequest({ includedTags: [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }] })),
            writer,
            context,
        );

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x69);
        expect(response.readUInt8()).toBe(6);
        expect(response.readUInt32LE()).toBe(0);

        const buffer = cachedTransfer(context);
        const reader = new BufferReader(buffer);
        expect(readTaggedPropertyValue(reader)).toEqual({ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hi" });
        expect(reader.hasMore()).toBe(false);
    });

    it("Honors an empty PropertyTags list literally (copies nothing), rather than falling back to a default set.", async () => {
        const context = makeContext({ messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }) } as any });
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const handler = new RopFastTransferSourceCopyPropertiesHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const buffer = cachedTransfer(context);
        expect(buffer.length).toBe(0);
    });

    it("Creates a fastTransfer output handle for a folder handle too.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const handler = new RopFastTransferSourceCopyPropertiesHandler();
        const writer = new BufferWriter();

        await handler.handle(
            new BufferReader(buildRequest({ includedTags: [{ propertyId: 0x3001, propertyType: PropertyType.PtypString }] })),
            writer,
            context,
        );

        expect(context.session.handles[6]?.type).toBe("fastTransfer");
        expect(context.session.handles[6]?.transferSourceType).toBe("folder");
    });

    it("Returns MAPI_E_TOO_BIG, creating no handle, for more than MAX_PROPERTY_TAG_COUNT tags.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "folder", entityUid: "folder:f1" };
        const tags = Array.from({ length: MAX_PROPERTY_TAG_COUNT + 1 }, () => ({ propertyId: 0x3001, propertyType: PropertyType.PtypString }));
        const writer = new BufferWriter();

        await new RopFastTransferSourceCopyPropertiesHandler().handle(new BufferReader(buildRequest({ includedTags: tags })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040305);
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[6]).toBeUndefined();
    });

    it("Returns MAPI_E_TOO_BIG, creating no handle, when the built stream exceeds MAX_FAST_TRANSFER_BYTES.", async () => {
        const subject = "a".repeat(MAX_FAST_TRANSFER_BYTES / 2 + 1); // UTF-16: two bytes per character
        const context = makeContext({ messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject }) } as any });
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const writer = new BufferWriter();

        await new RopFastTransferSourceCopyPropertiesHandler().handle(
            new BufferReader(buildRequest({ includedTags: [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }] })),
            writer,
            context,
        );

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040305);
        expect(context.session.handles[6]).toBeUndefined();
    });
});

function cachedTransfer(context: RopContext): Buffer {
    return handleDataCache.get(handleDataKey(context.session.uid, 6, context.session.handles[6].generation))!;
}
