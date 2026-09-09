///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, writePropertyTag } from "../../src/codec/PropertyValue.js";
import { RopSetColumnsHandler } from "../../src/rop/RopSetColumnsHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

function buildRequest(logonId: number, inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(0); // SetColumnsFlags
    writer.writeUInt16LE(tags.length);
    for (const tag of tags) {
        writePropertyTag(writer, tag);
    }
    return writer.toBuffer();
}

function makeContext(): RopContext {
    const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
    return { mailboxUid: "mailbox-1", userUid: "user-1", session, folderRepo: {} as any, messageRepo: {} as any, mailboxRepo: {} as any, folderClass: {} as any, messageClass: {} as any, scanPipeline: {} as any, mailTransport: {} as any, blobStore: {} as any };
}

describe("RopSetColumnsHandler Tests", () => {
    it("Has RopId 0x12.", () => {
        expect(new RopSetColumnsHandler().ropId).toBe(0x12);
    });

    it("Stores the configured columns on the referenced table handle and returns a well-formed response.", () => {
        const context = makeContext();
        context.session.handles[7] = { type: "table", entityUid: "folder:top1", rows: [], cursor: 0 };
        const handler = new RopSetColumnsHandler();
        const writer = new BufferWriter();
        const tags = [
            { propertyId: 0x3001, propertyType: PropertyType.PtypString },
            { propertyId: 0x6748, propertyType: PropertyType.PtypInteger64 },
        ];

        handler.handle(new BufferReader(buildRequest(0, 7, tags)), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x12);
        expect(response.readUInt8()).toBe(7); // echoes InputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt8()).toBe(0); // TableStatus - TBLSTAT_COMPLETE
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[7]?.columns).toEqual(tags);
    });

    it("Returns MAPI_E_INVALID_OBJECT when the referenced handle doesn't exist, instead of silently succeeding.", () => {
        const context = makeContext();
        const handler = new RopSetColumnsHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest(0, 42, [])), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x12);
        expect(response.readUInt8()).toBe(42);
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Returns MAPI_E_INVALID_OBJECT when the referenced handle exists but isn't a table (e.g. a message handle).", () => {
        const context = makeContext();
        context.session.handles[7] = { type: "message", entityUid: "message:m1" };
        const handler = new RopSetColumnsHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest(0, 7, [])), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(context.session.handles[7]?.columns).toBeUndefined();
    });
});
