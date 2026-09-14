///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType } from "../../src/codec/PropertyValue.js";
import { RopFastTransferSourceGetBufferHandler } from "../../src/rop/RopFastTransferSourceGetBufferHandler.js";
import { HandleDataCache, handleDataCache, handleDataKey, MemoryHandleDataStore } from "../../src/rop/HandleDataCache.js";
import { MAX_FAST_TRANSFER_BYTES } from "../../src/rop/FastTransferStream.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const BUFFER_SIZE_SERVER_DETERMINED = 0xbabe;

function buildRequest({ logonId = 0, inputHandleIndex = 5, bufferSize = 4096, maximumBufferSize }: { logonId?: number; inputHandleIndex?: number; bufferSize?: number; maximumBufferSize?: number }): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt16LE(bufferSize);
    if (bufferSize === BUFFER_SIZE_SERVER_DETERMINED) {
        writer.writeUInt16LE(maximumBufferSize ?? 0);
    }
    return writer.toBuffer();
}

function makeContext(): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
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
}

/** Opens a fastTransfer handle at index 5 whose built stream is already cached, the state CopyTo/CopyProperties leave. */
function seedTransfer(context: RopContext, payload: Buffer): void {
    context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferSourceType: "folder", transferPosition: 0, generation: "g7" };
    handleDataCache.set(handleDataKey(context.session.uid, 5, "g7"), payload);
}

/** Reads the fixed response header and returns [TransferStatus, chunk]. */
function readResponse(buffer: Buffer): [number, Buffer] {
    const response = new BufferReader(buffer);
    expect(response.readUInt8()).toBe(0x4e);
    expect(response.readUInt8()).toBe(5);
    expect(response.readUInt32LE()).toBe(0); // ReturnValue
    const status = response.readUInt16LE();
    expect(response.readUInt16LE()).toBe(0); // InProgressCount
    expect(response.readUInt16LE()).toBe(1); // TotalStepCount
    expect(response.readUInt8()).toBe(0); // Reserved
    const chunk = response.readBytes(response.readUInt16LE());
    expect(response.hasMore()).toBe(false);
    return [status, chunk];
}

describe("RopFastTransferSourceGetBufferHandler Tests", () => {
    it("Has RopId 0x4E.", () => {
        expect(new RopFastTransferSourceGetBufferHandler().ropId).toBe(0x4e);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a fastTransfer handle.", async () => {
        const context = makeContext();
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns the whole buffer in one call as Done when it fits within BufferSize.", async () => {
        const context = makeContext();
        const payload = Buffer.from("hello world");
        seedTransfer(context, payload);
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({ bufferSize: 4096 })), writer, context);

        const [status, chunk] = readResponse(writer.toBuffer());
        expect(status).toBe(0x0003); // Done
        expect(chunk).toEqual(payload);
        expect(context.session.handles[5]?.transferPosition).toBe(payload.length);
        // The stream lives in the cache, never in session state.
        expect(JSON.stringify(context.session)).not.toContain(payload.toString("base64"));
    });

    it("Pages the buffer across multiple calls, reporting Partial until the last chunk.", async () => {
        const context = makeContext();
        seedTransfer(context, Buffer.from("0123456789"));
        const handler = new RopFastTransferSourceGetBufferHandler();
        const results: [number, string][] = [];

        for (let i = 0; i < 3; i++) {
            const writer = new BufferWriter();
            await handler.handle(new BufferReader(buildRequest({ bufferSize: 4 })), writer, context);
            const [status, chunk] = readResponse(writer.toBuffer());
            results.push([status, chunk.toString()]);
        }

        expect(results).toEqual([
            [0x0001, "0123"],
            [0x0001, "4567"],
            [0x0003, "89"],
        ]);
        expect(context.session.handles[5]?.transferPosition).toBe(10);
    });

    it("Returns the whole remaining buffer in one call for the 0xBABE server-determined BufferSize sentinel.", async () => {
        const context = makeContext();
        seedTransfer(context, Buffer.from("a".repeat(500)));
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(
            new BufferReader(buildRequest({ bufferSize: BUFFER_SIZE_SERVER_DETERMINED, maximumBufferSize: 32768 })),
            writer,
            context,
        );

        const [status, chunk] = readResponse(writer.toBuffer());
        expect(status).toBe(0x0003);
        expect(chunk.length).toBe(500);
    });

    it("Honors MaximumBufferSize for the 0xBABE sentinel, and clamps to 0xFFFF, reporting Partial instead of throwing.", async () => {
        const context = makeContext();
        seedTransfer(context, Buffer.alloc(140000, 0x41));
        const handler = new RopFastTransferSourceGetBufferHandler();

        const limited = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ bufferSize: BUFFER_SIZE_SERVER_DETERMINED, maximumBufferSize: 32768 })), limited, context);
        const [limitedStatus, limitedChunk] = readResponse(limited.toBuffer());
        expect(limitedStatus).toBe(0x0001);
        expect(limitedChunk.length).toBe(32768);

        const clamped = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ bufferSize: 0xffff })), clamped, context);
        const [clampedStatus, clampedChunk] = readResponse(clamped.toBuffer());
        expect(clampedStatus).toBe(0x0001); // Partial - more remains past the 0xFFFF cap
        expect(clampedChunk.length).toBe(0xffff);
        expect(context.session.handles[5]?.transferPosition).toBe(32768 + 0xffff);
    });

    it("Holds the chunk to the room left in the ROP output buffer, reporting NoRoom when there is none.", async () => {
        const context = makeContext();
        seedTransfer(context, Buffer.alloc(1000, 0x42));
        const handler = new RopFastTransferSourceGetBufferHandler();

        context.ropOutputRemaining = 15 + 100;
        const partial = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ bufferSize: 4096 })), partial, context);
        const [partialStatus, partialChunk] = readResponse(partial.toBuffer());
        expect(partial.toBuffer().length).toBe(115);
        expect(partialStatus).toBe(0x0001);
        expect(partialChunk.length).toBe(100);

        context.ropOutputRemaining = 10;
        const noRoom = new BufferWriter();
        await handler.handle(new BufferReader(buildRequest({ bufferSize: 4096 })), noRoom, context);
        const [noRoomStatus, noRoomChunk] = readResponse(noRoom.toBuffer());
        expect(noRoomStatus).toBe(0x0002); // NoRoom
        expect(noRoomChunk.length).toBe(0);
        expect(context.session.handles[5]?.transferPosition).toBe(100);
    });

    it("Fails with a transfer error instead of serving bytes of a rebuilt stream when the stream is lost part-way through.", async () => {
        const context = makeContext();
        const findOne = vi.fn().mockResolvedValue({ uid: "m1", subject: "Changed since" });
        context.messageRepo = { findOne } as any;
        context.session.handles[5] = { type: "fastTransfer", entityUid: "message:m1", transferSourceType: "message", transferPosition: 40, generation: "lost-one" };
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80004005);
        expect(response.hasMore()).toBe(false);
        expect(findOne).not.toHaveBeenCalled();
        expect(context.session.handles[5].transferPosition).toBe(40);
    });

    it("Pages a stream another replica built from shared handle data storage.", async () => {
        const context = makeContext();
        const shared = new MemoryHandleDataStore(new HandleDataCache());
        context.handleData = shared;
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferSourceType: "folder", transferPosition: 3, generation: "shared" };
        await shared.set(handleDataKey(context.session.uid, 5, "shared"), Buffer.from("0123456789"), 60);
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({ bufferSize: 4 })), writer, context);

        const [, chunk] = readResponse(writer.toBuffer());
        expect(chunk.toString()).toBe("3456");
    });

    it("Fails a rebuild that passes MAX_FAST_TRANSFER_BYTES with MAPI_E_TOO_BIG.", async () => {
        const context = makeContext();
        context.session.handles[5] = { type: "fastTransfer", entityUid: "message:m1", transferSourceType: "message", transferPosition: 0, generation: "too-big" };
        const built = vi.spyOn(BufferWriter.prototype, "length", "get").mockReturnValue(MAX_FAST_TRANSFER_BYTES + 1);
        context.messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Huge" }) } as any;
        const writer = new BufferWriter();
        try {
            await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({})), writer, context);
        } finally {
            built.mockRestore();
        }

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040305);
    });

    it("Rebuilds the stream from the handle's recorded source on a cache miss (e.g. another replica built it).", async () => {
        const context = makeContext();
        context.messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Rebuilt" }) } as any;
        context.session.handles[5] = {
            type: "fastTransfer",
            entityUid: "message:m1",
            transferSourceType: "message",
            transferColumns: [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }],
            transferPosition: 0,
            generation: "g99",
        };
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const [status, chunk] = readResponse(writer.toBuffer());
        expect(status).toBe(0x0003);
        expect(chunk.includes(Buffer.from("Rebuilt", "utf16le"))).toBe(true);
        expect(handleDataCache.get(handleDataKey(context.session.uid, 5, "g99"))).toEqual(chunk);
    });

    it("Rebuilds a CopyTo-style handle (exclude list, no explicit columns) with the excluded property left out.", async () => {
        const context = makeContext();
        context.messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Skipped" }) } as any;
        context.session.handles[5] = { type: "fastTransfer", entityUid: "message:m1", transferSourceType: "message", transferExcludeIds: [0x0037], generation: "g100" };
        const writer = new BufferWriter();

        await new RopFastTransferSourceGetBufferHandler().handle(new BufferReader(buildRequest({})), writer, context);

        const [, chunk] = readResponse(writer.toBuffer());
        expect(chunk.length).toBeGreaterThan(0);
        expect(chunk.includes(Buffer.from("Skipped", "utf16le"))).toBe(false);
    });
});
