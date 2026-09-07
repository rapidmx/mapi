///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopReadStreamHandler } from "../../src/rop/RopReadStreamHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { InMemoryBlobStore } from "../testDoubles.js";

function buildRequest({ logonId = 0, inputHandleIndex = 6, byteCount = 100, maximumByteCount }: { logonId?: number; inputHandleIndex?: number; byteCount?: number; maximumByteCount?: number }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt16LE(byteCount);
    if (byteCount === 0xbabe) {
        writer.writeUInt32LE(maximumByteCount ?? 0);
    }
    return writer.toBuffer();
}

const BODY_TEXT = "Hello body.";

async function makeContext(): Promise<RopContext> {
    const blobStore = new InMemoryBlobStore();
    await blobStore.put("bodies/m1", Buffer.from(`Subject: Hi\r\n\r\n${BODY_TEXT}`));
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    session.handles[6] = { type: "stream", entityUid: "message:m1", propertyId: 0x1000, propertyType: 0x001f, streamPosition: 0 };
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session,
        folderRepo: {} as any,
        messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", bodyBlobKey: "bodies/m1" }) } as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore,
    };
}

describe("RopReadStreamHandler Tests", () => {
    it("Has RopId 0x2C.", () => {
        expect(new RopReadStreamHandler().ropId).toBe(0x2c);
    });

    it("Reads the whole stream in one call when ByteCount exceeds StreamSize, advancing streamPosition.", async () => {
        const context = await makeContext();
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ byteCount: 1000 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x2c);
        expect(response.readUInt8()).toBe(6);
        expect(response.readUInt32LE()).toBe(0);
        const dataSize = response.readUInt16LE();
        const expectedBytes = Buffer.concat([Buffer.from(BODY_TEXT, "utf16le"), Buffer.from([0, 0])]);
        expect(dataSize).toBe(expectedBytes.length);
        expect(response.readBytes(dataSize)).toEqual(expectedBytes);
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[6]?.streamPosition).toBe(expectedBytes.length);
    });

    it("Respects ByteCount, leaving the remainder for a subsequent call.", async () => {
        const context = await makeContext();
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ byteCount: 4 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(4);
        expect(context.session.handles[6]?.streamPosition).toBe(4);
    });

    it("Uses MaximumByteCount when ByteCount is the 0xBABE sentinel.", async () => {
        const context = await makeContext();
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ byteCount: 0xbabe, maximumByteCount: 6 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(6);
    });

    it("Returns MAPI_E_INVALID_OBJECT (with DataSize 0) when InputHandleIndex isn't a stream handle.", async () => {
        const context = await makeContext();
        context.session.handles[6] = { type: "message", entityUid: "message:m1" };
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.readUInt16LE()).toBe(0);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex references no handle at all.", async () => {
        const context = await makeContext();
        delete context.session.handles[6];
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
    });

    it("Defaults streamPosition to 0 when a stream handle was constructed without one.", async () => {
        const context = await makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "message:m1", propertyId: 0x1000, propertyType: 0x001f };
        const handler = new RopReadStreamHandler();
        const writer = new BufferWriter();

        await handler.handle(new BufferReader(buildRequest({ byteCount: 4 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(4);
        expect(context.session.handles[6]?.streamPosition).toBe(4);
    });
});
