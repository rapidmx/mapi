///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopFastTransferSourceGetBufferHandler } from "../../src/rop/RopFastTransferSourceGetBufferHandler.js";
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

describe("RopFastTransferSourceGetBufferHandler Tests", () => {
    it("Has RopId 0x4E.", () => {
        expect(new RopFastTransferSourceGetBufferHandler().ropId).toBe(0x4e);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a fastTransfer handle.", () => {
        const context = makeContext();
        const handler = new RopFastTransferSourceGetBufferHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns the whole buffer in one call as Done when it fits within BufferSize.", () => {
        const context = makeContext();
        const payload = Buffer.from("hello world");
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferBufferBase64: payload.toString("base64"), transferPosition: 0 };
        const handler = new RopFastTransferSourceGetBufferHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({ bufferSize: 4096 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x4e);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt16LE()).toBe(0x0003); // TransferStatus - Done
        expect(response.readUInt16LE()).toBe(0); // InProgressCount
        expect(response.readUInt16LE()).toBe(1); // TotalStepCount
        expect(response.readUInt8()).toBe(0); // Reserved
        const size = response.readUInt16LE();
        expect(size).toBe(payload.length);
        expect(response.readBytes(size)).toEqual(payload);
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[5]?.transferPosition).toBe(payload.length);
    });

    it("Pages the buffer across multiple calls, reporting Partial until the last chunk.", () => {
        const context = makeContext();
        const payload = Buffer.from("0123456789");
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferBufferBase64: payload.toString("base64"), transferPosition: 0 };
        const handler = new RopFastTransferSourceGetBufferHandler();

        const writer1 = new BufferWriter();
        handler.handle(new BufferReader(buildRequest({ bufferSize: 4 })), writer1, context);
        const response1 = new BufferReader(writer1.toBuffer());
        response1.readUInt8();
        response1.readUInt8();
        response1.readUInt32LE();
        expect(response1.readUInt16LE()).toBe(0x0001); // Partial
        response1.readUInt16LE();
        response1.readUInt16LE();
        response1.readUInt8();
        const size1 = response1.readUInt16LE();
        expect(size1).toBe(4);
        expect(response1.readBytes(size1).toString()).toBe("0123");
        expect(context.session.handles[5]?.transferPosition).toBe(4);

        const writer2 = new BufferWriter();
        handler.handle(new BufferReader(buildRequest({ bufferSize: 4 })), writer2, context);
        const response2 = new BufferReader(writer2.toBuffer());
        response2.readUInt8();
        response2.readUInt8();
        response2.readUInt32LE();
        expect(response2.readUInt16LE()).toBe(0x0001); // Partial - 4 bytes returned, 2 remain
        response2.readUInt16LE();
        response2.readUInt16LE();
        response2.readUInt8();
        const size2 = response2.readUInt16LE();
        expect(response2.readBytes(size2).toString()).toBe("4567");

        const writer3 = new BufferWriter();
        handler.handle(new BufferReader(buildRequest({ bufferSize: 4 })), writer3, context);
        const response3 = new BufferReader(writer3.toBuffer());
        response3.readUInt8();
        response3.readUInt8();
        response3.readUInt32LE();
        expect(response3.readUInt16LE()).toBe(0x0003); // Done - final 2 bytes
        response3.readUInt16LE();
        response3.readUInt16LE();
        response3.readUInt8();
        const size3 = response3.readUInt16LE();
        expect(response3.readBytes(size3).toString()).toBe("89");
        expect(context.session.handles[5]?.transferPosition).toBe(10);
    });

    it("Returns the whole remaining buffer in one call for the 0xBABE server-determined BufferSize sentinel.", () => {
        const context = makeContext();
        const payload = Buffer.from("a".repeat(500));
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferBufferBase64: payload.toString("base64"), transferPosition: 0 };
        const handler = new RopFastTransferSourceGetBufferHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({ bufferSize: BUFFER_SIZE_SERVER_DETERMINED, maximumBufferSize: 32768 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(0x0003); // Done
        response.readUInt16LE();
        response.readUInt16LE();
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(500);
    });

    it("Clamps TransferBufferSize to 0xFFFF for the 0xBABE sentinel when the buffer exceeds a uint16, reporting Partial instead of throwing.", () => {
        const context = makeContext();
        const payload = Buffer.alloc(70000, 0x41); // far larger than 0xFFFF bytes
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1", transferBufferBase64: payload.toString("base64"), transferPosition: 0 };
        const handler = new RopFastTransferSourceGetBufferHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({ bufferSize: BUFFER_SIZE_SERVER_DETERMINED, maximumBufferSize: 32768 })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - success, not a RangeError
        expect(response.readUInt16LE()).toBe(0x0001); // Partial - more remains past the 0xFFFF cap
        response.readUInt16LE();
        response.readUInt16LE();
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(0xffff);
        expect(context.session.handles[5]?.transferPosition).toBe(0xffff);
    });

    it("Treats an absent transferBufferBase64 as an empty buffer, reporting Done with a zero-length chunk.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "fastTransfer", entityUid: "folder:f1" };
        const handler = new RopFastTransferSourceGetBufferHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(0x0003); // Done
        response.readUInt16LE();
        response.readUInt16LE();
        response.readUInt8();
        expect(response.readUInt16LE()).toBe(0);
    });
});
