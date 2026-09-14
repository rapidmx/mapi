///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Round-3 review limits: request framing, tag-array and ROP-count caps, write-stream growth, table paging and
// handle release semantics. Pure logic, no DB.
import { decodeExecuteRequest, MAX_ROP_BUFFER_SIZE } from "../src/BaseMapiEmsmdbRoute.js";
import { BufferReader, BufferWriter } from "../src/codec/BufferCursor.js";
import { MAX_PROPERTY_TAG_COUNT, PropertyType, readPropertyTagArray, writePropertyTag } from "../src/codec/PropertyValue.js";
import { decodeRopBuffer, encodeRopBuffer, MAX_HANDLE_TABLE_ENTRIES } from "../src/codec/RopBuffer.js";
import { MapiSessionContext, assignHandle, releaseHandle } from "../src/MapiSessionManager.js";
import { readLargePropertyTagArray } from "../src/nspi/NspiCodec.js";
import { dispatchRops, MAX_ROPS_PER_EXECUTE } from "../src/RopDispatcher.js";
import { handleDataCache, handleDataKey } from "../src/rop/HandleDataCache.js";
import type { RopContext, RopHandler } from "../src/rop/RopHandler.js";
import { RopGetPropertiesSpecificHandler } from "../src/rop/RopGetPropertiesSpecificHandler.js";
import { MAX_ROWS_PER_QUERY, RopQueryRowsHandler } from "../src/rop/RopQueryRowsHandler.js";
import { RopReleaseHandler } from "../src/rop/RopReleaseHandler.js";
import { RopSetColumnsHandler } from "../src/rop/RopSetColumnsHandler.js";
import { MAX_WRITE_STREAM_BYTES, readWriteStream, RopWriteStreamHandler } from "../src/rop/RopWriteStreamHandler.js";

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
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
        ...overrides,
    };
}

function tags(count: number): Buffer {
    const writer = new BufferWriter();
    for (let i = 0; i < count; i++) {
        writePropertyTag(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString });
    }
    return writer.toBuffer();
}

function executeBody(ropBufferSize: number, ropBuffer: Buffer): Buffer {
    return new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBufferSize).writeBytes(ropBuffer).writeUInt32LE(0).writeUInt32LE(0).toBuffer();
}

describe("Execute request framing", () => {
    it("Decodes a well-formed Execute body's ROP buffer and handle table.", () => {
        const ropBuffer = encodeRopBuffer({ ropsList: Buffer.from([0x01, 0x00, 0x02]), handleTable: [7, 8] });
        const decoded = decodeExecuteRequest(executeBody(ropBuffer.length, ropBuffer));
        expect(decoded.ropsList).toEqual(Buffer.from([0x01, 0x00, 0x02]));
        expect(decoded.handleTable).toEqual([7, 8]);
    });

    it("Rejects a RopBufferSize over MAX_ROP_BUFFER_SIZE with a 400, even when the body really is that long.", () => {
        const ropBuffer = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
        const padded = Buffer.concat([ropBuffer, Buffer.alloc(MAX_ROP_BUFFER_SIZE)]);
        expect(() => decodeExecuteRequest(executeBody(MAX_ROP_BUFFER_SIZE + 1, padded))).toThrow(expect.objectContaining({ status: 400 }));
    });

    it("Rejects a RopBufferSize past the end of the body, and a truncated body, with a 400 instead of a 500.", () => {
        const body = new BufferWriter().writeUInt32LE(0).writeUInt32LE(100).writeBytes(Buffer.alloc(10)).toBuffer();
        expect(() => decodeExecuteRequest(body)).toThrow(expect.objectContaining({ status: 400 }));
        expect(() => decodeExecuteRequest(Buffer.alloc(3))).toThrow(expect.objectContaining({ status: 400 }));
    });

    it("Rejects a handle table with more than MAX_HANDLE_TABLE_ENTRIES entries or a partial entry.", () => {
        const tooMany = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: Array.from({ length: MAX_HANDLE_TABLE_ENTRIES + 1 }, (_, i) => i) });
        expect(() => decodeRopBuffer(tooMany)).toThrow(RangeError);
        expect(() => decodeExecuteRequest(executeBody(tooMany.length, tooMany))).toThrow(expect.objectContaining({ status: 400 }));

        const atLimit = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: Array.from({ length: MAX_HANDLE_TABLE_ENTRIES }, (_, i) => i) });
        expect(decodeRopBuffer(atLimit).handleTable.length).toBe(MAX_HANDLE_TABLE_ENTRIES);

        const partial = Buffer.concat([encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [1] }), Buffer.from([0xff])]);
        expect(() => decodeRopBuffer(partial)).toThrow(RangeError);
    });
});

describe("ROP count per Execute", () => {
    it("Stops dispatching after MAX_ROPS_PER_EXECUTE ROPs.", async () => {
        const context = makeContext();
        const handled = vi.fn((reader: BufferReader) => {
            reader.readUInt8();
        });
        const handlers = new Map<number, RopHandler>([[0x01, { ropId: 0x01, handle: handled }]]);
        const ropsList = Buffer.alloc((MAX_ROPS_PER_EXECUTE + 10) * 2);
        for (let i = 0; i < MAX_ROPS_PER_EXECUTE + 10; i++) {
            ropsList[i * 2] = 0x01;
        }

        await dispatchRops(ropsList, handlers, context);

        expect(handled).toHaveBeenCalledTimes(MAX_ROPS_PER_EXECUTE);
    });
});

describe("Property tag arrays", () => {
    it("Reads up to MAX_PROPERTY_TAG_COUNT tags, and skips (returning undefined) a larger array.", () => {
        expect(readPropertyTagArray(new BufferReader(tags(MAX_PROPERTY_TAG_COUNT)), MAX_PROPERTY_TAG_COUNT)).toHaveLength(MAX_PROPERTY_TAG_COUNT);

        const reader = new BufferReader(Buffer.concat([tags(MAX_PROPERTY_TAG_COUNT + 1), Buffer.from([0xaa])]));
        expect(readPropertyTagArray(reader, MAX_PROPERTY_TAG_COUNT + 1)).toBeUndefined();
        expect(reader.readUInt8()).toBe(0xaa); // positioned just past the skipped array
    });

    it("Throws when a tag count claims more bytes than the buffer holds.", () => {
        expect(() => readPropertyTagArray(new BufferReader(tags(2)), 60000)).toThrow(RangeError);
    });

    it("Caps an NSPI LargePropertyTagArray the same way.", () => {
        const body = (count: number) => Buffer.concat([new BufferWriter().writeUInt32LE(count).toBuffer(), tags(count)]);
        expect(readLargePropertyTagArray(new BufferReader(body(MAX_PROPERTY_TAG_COUNT)))).toHaveLength(MAX_PROPERTY_TAG_COUNT);
        expect(() => readLargePropertyTagArray(new BufferReader(body(MAX_PROPERTY_TAG_COUNT + 1)))).toThrow(RangeError);
        expect(() => readLargePropertyTagArray(new BufferReader(new BufferWriter().writeUInt32LE(0xffffffff).toBuffer()))).toThrow(RangeError);
    });

    it("RopSetColumns answers MAPI_E_TOO_BIG for too many columns and leaves the table's columns alone.", () => {
        const context = makeContext();
        context.session.handles[3] = { type: "table", entityUid: "folder:f1", rows: [], columns: [{ propertyId: 1, propertyType: 3 }] };
        const request = Buffer.concat([new BufferWriter().writeUInt8(0).writeUInt8(3).writeUInt8(0).writeUInt16LE(MAX_PROPERTY_TAG_COUNT + 1).toBuffer(), tags(MAX_PROPERTY_TAG_COUNT + 1)]);
        const writer = new BufferWriter();

        new RopSetColumnsHandler().handle(new BufferReader(request), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x12);
        expect(response.readUInt8()).toBe(3);
        expect(response.readUInt32LE()).toBe(0x80040305);
        expect(response.hasMore()).toBe(false);
        expect(context.session.handles[3].columns).toEqual([{ propertyId: 1, propertyType: 3 }]);
    });

    it("RopGetPropertiesSpecific answers MAPI_E_TOO_BIG for too many tags without resolving anything.", async () => {
        const findOne = vi.fn();
        const context = makeContext({ messageRepo: { findOne } as any });
        context.session.handles[3] = { type: "message", entityUid: "message:m1" };
        const header = new BufferWriter().writeUInt8(0).writeUInt8(3).writeUInt16LE(0).writeUInt16LE(1).writeUInt16LE(MAX_PROPERTY_TAG_COUNT + 1).toBuffer();
        const writer = new BufferWriter();

        await new RopGetPropertiesSpecificHandler().handle(new BufferReader(Buffer.concat([header, tags(MAX_PROPERTY_TAG_COUNT + 1)])), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80040305);
        expect(findOne).not.toHaveBeenCalled();
    });
});

describe("Write stream growth", () => {
    it("Tracks writeSize and refuses a write past MAX_WRITE_STREAM_BYTES with MAPI_E_TOO_BIG, storing nothing.", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "g6", writeTargetHandleIndex: 3, writeSize: MAX_WRITE_STREAM_BYTES - 2 };
        const handler = new RopWriteStreamHandler();
        const request = (data: Buffer) => new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(data.length).writeBytes(data).toBuffer();

        const ok = new BufferWriter();
        await handler.handle(new BufferReader(request(Buffer.from("ab"))), ok, context);
        const okResponse = new BufferReader(ok.toBuffer());
        okResponse.readUInt8();
        okResponse.readUInt8();
        expect(okResponse.readUInt32LE()).toBe(0);
        expect(context.session.handles[6].writeSize).toBe(MAX_WRITE_STREAM_BYTES);

        const refused = new BufferWriter();
        await handler.handle(new BufferReader(request(Buffer.from("c"))), refused, context);
        const refusedResponse = new BufferReader(refused.toBuffer());
        refusedResponse.readUInt8();
        refusedResponse.readUInt8();
        expect(refusedResponse.readUInt32LE()).toBe(0x80040305);
        expect(refusedResponse.readUInt16LE()).toBe(0);
        expect(context.session.handles[6].writeSize).toBe(MAX_WRITE_STREAM_BYTES);
    });

    it("Keeps written bytes out of the session, and a retried write at the same offset replaces its chunk.", async () => {
        const context = makeContext();
        context.session.handles[6] = { type: "stream", entityUid: "", generation: "g-retry", writeTargetHandleIndex: 3, writeSize: 0 };
        const handler = new RopWriteStreamHandler();
        const write = (text: string) =>
            handler.handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(text.length).writeBytes(Buffer.from(text)).toBuffer()), new BufferWriter(), context);

        await write("hello ");
        const savedBeforeFailedRequest = JSON.stringify(context.session);
        await write("wrong"); // this request's session save is lost...
        context.session.handles[6].writeSize = 6; // ...so the stored session still says 6 bytes
        await write("world");
        await write(""); // an empty write stores nothing

        expect(savedBeforeFailedRequest).not.toContain(Buffer.from("hello ").toString("base64"));
        expect((await readWriteStream(context, 6, context.session.handles[6]))!.toString()).toBe("hello world");
    });
});

describe("RopQueryRows paging", () => {
    function queryRows(context: RopContext, count: number): Promise<number> {
        const writer = new BufferWriter();
        const request = new BufferWriter().writeUInt8(0).writeUInt8(4).writeUInt8(0).writeUInt8(1).writeUInt16LE(count).toBuffer();
        return new RopQueryRowsHandler().handle(new BufferReader(request), writer, context).then(() => {
            const response = new BufferReader(writer.toBuffer());
            response.readUInt8();
            response.readUInt8();
            response.readUInt32LE();
            response.readUInt8();
            return response.readUInt16LE();
        });
    }

    it("Reads a contents table's rows from the database window by window, advancing the cursor.", async () => {
        const find = vi.fn().mockImplementation((_query: any, options: { page: number; limit: number }) =>
            Promise.resolve(Array.from({ length: Math.max(0, Math.min(options.limit, 7 - options.page * options.limit)) }, (_, i) => ({ uid: `m${i}` }))),
        );
        const context = makeContext({ messageRepo: { find, findOne: vi.fn().mockResolvedValue(undefined) } as any });
        context.session.handles[4] = { type: "table", entityUid: "folder:f1", contentsKind: "message", cursor: 0, columns: [] };

        expect(await queryRows(context, 5)).toBe(5);
        expect(context.session.handles[4].cursor).toBe(5);
        expect(await queryRows(context, 5)).toBe(2);
        expect(context.session.handles[4].cursor).toBe(7);
        expect(context.session.handles[4].rows).toBeUndefined();
    });

    it("Returns at most MAX_ROWS_PER_QUERY rows per call.", async () => {
        const context = makeContext();
        context.session.handles[4] = { type: "table", entityUid: "folder:f1", rows: Array.from({ length: 2000 }, (_, i) => `virtual:v${i}`), cursor: 0, columns: [] };

        expect(await queryRows(context, 0xffff)).toBe(MAX_ROWS_PER_QUERY);
        expect(context.session.handles[4].cursor).toBe(MAX_ROWS_PER_QUERY);
    });
});

describe("Handle generations and release", () => {
    it("Gives every assignment a new generation, even at a reused index.", () => {
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        const first = assignHandle(session, 5, { type: "message", entityUid: "" });
        const second = assignHandle(session, 5, { type: "message", entityUid: "" });
        expect(second.generation).toEqual(expect.any(String));
        expect(second.generation).not.toBe(first.generation);
        expect(session.handles[5]).toBe(second);
    });

    it("RopRelease releases write streams opened against the released message and drops cached handle data.", () => {
        const context = makeContext();
        const message = assignHandle(context.session, 5, { type: "message", entityUid: "" });
        const stream = assignHandle(context.session, 6, { type: "stream", entityUid: "", writeTargetHandleIndex: 5, writeTargetGeneration: message.generation });
        const unrelated = assignHandle(context.session, 7, { type: "stream", entityUid: "", writeTargetHandleIndex: 5, writeTargetGeneration: "another-generation" });
        handleDataCache.set(handleDataKey(context.session.uid, 6, stream.generation), Buffer.from("cached"));

        new RopReleaseHandler().handle(new BufferReader(Buffer.from([0x00, 0x05])), new BufferWriter(), context);

        expect(context.session.handles[5]).toBeUndefined();
        expect(context.session.handles[6]).toBeUndefined();
        expect(context.session.handles[7]).toBe(unrelated);
        expect(handleDataCache.get(handleDataKey(context.session.uid, 6, stream.generation))).toBeUndefined();
    });

    it("Reusing a message handle's index releases the old message's write streams too.", () => {
        const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
        const message = assignHandle(session, 5, { type: "message", entityUid: "" });
        assignHandle(session, 6, { type: "stream", entityUid: "", writeTargetHandleIndex: 5, writeTargetGeneration: message.generation });

        assignHandle(session, 5, { type: "message", entityUid: "" });

        expect(session.handles[6]).toBeUndefined();
        expect(() => releaseHandle(session, 42)).not.toThrow();
    });
});
