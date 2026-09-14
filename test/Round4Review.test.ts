///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Round-4 review fixes that live below the route: output-space clamping, the per-Execute work budget, MID eviction,
// and recipient parsing. Pure logic, no DB.
import { FolderType } from "@rapidmx/restapi";
import { decodeExecuteRequest } from "../src/BaseMapiEmsmdbRoute.js";
import { BufferReader, BufferWriter } from "../src/codec/BufferCursor.js";
import { PropertyType } from "../src/codec/PropertyValue.js";
import { encodeRopBuffer } from "../src/codec/RopBuffer.js";
import { MapiSessionContext } from "../src/MapiSessionManager.js";
import { dispatchRops, MAX_ROPS_LIST_BYTES } from "../src/RopDispatcher.js";
import { isPlainEmailAddress, parseAddressList, parseRecipientEntry, parseRecipientList, resolveRecipientList, splitAddressList } from "../src/rop/AddressList.js";
import { ExecuteBudget, MAX_BYTES_BUILT_PER_EXECUTE, MAX_ROWS_RESOLVED_PER_EXECUTE, WorkBudgetExceededError } from "../src/rop/ExecuteBudget.js";
import { buildFastTransferStream, FastTransferTooBigError, openFastTransferHandle } from "../src/rop/FastTransferStream.js";
import { HandleDataCache, handleDataKey, MemoryHandleDataStore } from "../src/rop/HandleDataCache.js";
import { loadStreamBody } from "../src/rop/MessageBodyStream.js";
import { assignOrGetMid, MAX_MESSAGE_IDS } from "../src/rop/MessageTarget.js";
import { resolvePropertyValues } from "../src/rop/PropertyResolvers.js";
import type { RopContext, RopHandler } from "../src/rop/RopHandler.js";
import { RopQueryRowsHandler } from "../src/rop/RopQueryRowsHandler.js";
import { RopReadStreamHandler } from "../src/rop/RopReadStreamHandler.js";
import { InMemoryBlobStore } from "./testDoubles.js";

function makeContext(overrides: Partial<RopContext> = {}): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: { findOne: vi.fn().mockResolvedValue(undefined), find: vi.fn().mockResolvedValue([]) } as any,
        messageRepo: { find: vi.fn().mockResolvedValue([]), findOne: vi.fn().mockResolvedValue(undefined) } as any,
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

describe("Execute MaxRopOut", () => {
    it("Reads MaxRopOut, defaulting to the largest RopBuffer when the body ends before it.", () => {
        const ropBuffer = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
        const withMax = new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).writeUInt32LE(4096).toBuffer();
        const withoutMax = new BufferWriter().writeUInt32LE(0).writeUInt32LE(ropBuffer.length).writeBytes(ropBuffer).toBuffer();

        expect(decodeExecuteRequest(withMax).maxRopOut).toBe(4096);
        expect(decodeExecuteRequest(withoutMax).maxRopOut).toBe(MAX_ROPS_LIST_BYTES + 2);
    });
});

describe("RopReadStream output space", () => {
    it("Returns no more data than fits in the room left in the ROP output buffer.", async () => {
        const blobStore = new InMemoryBlobStore();
        await blobStore.put("bodies/m1", Buffer.from(`Subject: Hi\r\n\r\n${"x".repeat(500)}`));
        const context = makeContext({ blobStore, messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", bodyBlobKey: "bodies/m1" }) } as any });
        context.session.handles[6] = { type: "stream", entityUid: "message:m1", streamPosition: 0, generation: "read-room" };
        context.ropOutputRemaining = 8 + 30;
        const writer = new BufferWriter();

        await new RopReadStreamHandler().handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(0xbabe).writeUInt32LE(100000).toBuffer()), writer, context);

        expect(writer.length).toBe(38);
        expect(writer.toBuffer().readUInt16LE(6)).toBe(30);
        expect(context.session.handles[6].streamPosition).toBe(30);

        context.ropOutputRemaining = 3;
        const none = new BufferWriter();
        await new RopReadStreamHandler().handle(new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(6).writeUInt16LE(10).toBuffer()), none, context);
        expect(none.toBuffer().readUInt16LE(6)).toBe(0);
        expect(context.session.handles[6].streamPosition).toBe(30);
    });
});

describe("RopQueryRows output space", () => {
    const request = (count: number) => new BufferReader(new BufferWriter().writeUInt8(0).writeUInt8(4).writeUInt8(0).writeUInt8(1).writeUInt16LE(count).toBuffer());
    const columns = [{ propertyId: 0x3001, propertyType: PropertyType.PtypString }];

    it("Returns only the rows that fit, advancing the cursor past just those.", async () => {
        const context = makeContext();
        context.session.handles[4] = { type: "table", entityUid: "virtual:root", rows: ["virtual:root", "virtual:search", "virtual:views"], cursor: 0, columns };
        // One row is Flags (1) + the UTF-16 display name + terminator: "Root" is 11 bytes, "Finder" 15.
        context.ropOutputRemaining = 9 + 11 + 15;
        const writer = new BufferWriter();

        await new RopQueryRowsHandler().handle(request(10), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readBytes(7);
        expect(response.readUInt16LE()).toBe(2);
        expect(context.session.handles[4].cursor).toBe(2);
    });

    it("Answers ecBufferTooSmall, leaving the cursor, when not even one row fits.", async () => {
        const context = makeContext();
        context.session.handles[4] = { type: "table", entityUid: "virtual:root", rows: ["virtual:root"], cursor: 0, columns };
        context.ropOutputRemaining = 12;
        const writer = new BufferWriter();

        await new RopQueryRowsHandler().handle(request(10), writer, context);

        expect(writer.toBuffer()).toEqual(Buffer.from([0x15, 4, 0x7d, 0x04, 0, 0]));
        expect(context.session.handles[4].cursor).toBe(0);
    });
});

describe("MID map cap", () => {
    it("Forgets the oldest MIDs past MAX_MESSAGE_IDS, and a forgotten item gets a new MID when seen again.", () => {
        const session = new MapiSessionContext({ mailboxUid: "m", userUid: "u" });
        for (let i = 0; i < MAX_MESSAGE_IDS + 5; i++) {
            assignOrGetMid(session, `message:${i}`);
        }

        expect(Object.keys(session.messageIds)).toHaveLength(MAX_MESSAGE_IDS);
        expect(Object.keys(session.messageTargetIds)).toHaveLength(MAX_MESSAGE_IDS);
        expect(session.messageIds["5"]).toBeUndefined();
        expect(session.messageIds["6"]).toBe("message:5");
        expect(assignOrGetMid(session, "message:6")).toBe(7);
        expect(assignOrGetMid(session, "message:0")).toBe(MAX_MESSAGE_IDS + 6);
    });
});

describe("Per-Execute work budget", () => {
    it("Throws once rows or bytes are used up.", () => {
        const budget = new ExecuteBudget(2, 10);
        budget.chargeRows();
        budget.chargeRows();
        expect(() => budget.chargeRows()).toThrow(WorkBudgetExceededError);
        budget.chargeBytes(10);
        expect(() => budget.chargeBytes(1)).toThrow(/byte budget/);
        expect(new ExecuteBudget().rowsRemaining).toBe(MAX_ROWS_RESOLVED_PER_EXECUTE);
        expect(new ExecuteBudget().bytesRemaining).toBe(MAX_BYTES_BUILT_PER_EXECUTE);
    });

    it("Counts every row resolved, so a chain of ROPs in one Execute fails once it passes the budget.", async () => {
        const context = makeContext({ budget: new ExecuteBudget(3) });
        const resolve = () => resolvePropertyValues("message:m1", [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }], context);
        await resolve();
        await resolve();
        await resolve();
        await expect(resolve()).rejects.toBeInstanceOf(WorkBudgetExceededError);
    });

    it("Fails a ROP that passes the budget with MAPI_E_TOO_COMPLEX through the dispatcher, keeping the rest of the request.", async () => {
        const handler: RopHandler = {
            ropId: 0x50,
            handle: async (reader, writer, context) => {
                reader.readUInt8();
                reader.readUInt8();
                await resolvePropertyValues("message:m1", [], context);
                writer.writeUInt8(0x50).writeUInt8(0).writeUInt32LE(0);
            },
        };
        const context = makeContext({ budget: new ExecuteBudget(1) });

        const response = await dispatchRops(Buffer.from([0x50, 0, 0, 0x50, 0, 1]), new Map([[0x50, handler]]), context);

        expect(response).toEqual(Buffer.from([0x50, 0, 0, 0, 0, 0, 0x50, 1, 0x17, 0x01, 0x04, 0x80]));
    });

    it("Stops building a FastTransfer stream as soon as it passes its byte limit, instead of building it whole first.", async () => {
        const findOne = vi.fn().mockResolvedValue({ uid: "m", subject: "A fairly long subject line" });
        const context = makeContext({
            folderRepo: { findOne: vi.fn().mockResolvedValue({ uid: "f1", type: FolderType.INBOX }), find: vi.fn().mockResolvedValue([]) } as any,
            messageRepo: { find: vi.fn().mockResolvedValue(Array.from({ length: 50 }, (_, i) => ({ uid: `m${i}` }))), findOne } as any,
            mailboxRepo: { findOne: vi.fn().mockResolvedValue(undefined) } as any,
        });

        await expect(buildFastTransferStream({ type: "folder", entityUid: "folder:f1" }, context, { maxBytes: 200 })).rejects.toBeInstanceOf(FastTransferTooBigError);
        expect(findOne.mock.calls.length).toBeLessThan(10);
    });

    it("Counts a built stream's bytes, and stores it in the context's handle data store.", async () => {
        const handleData = new MemoryHandleDataStore(new HandleDataCache());
        const budget = new ExecuteBudget();
        const context = makeContext({ budget, handleData, messageRepo: { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }) } as any });

        expect(await openFastTransferHandle(context, 6, { type: "message", entityUid: "message:m1" }, {})).toBe(true);

        const stored = await handleData.get(handleDataKey(context.session.uid, 6, context.session.handles[6].generation));
        expect(stored!.length).toBeGreaterThan(0);
        expect(budget.bytesRemaining).toBe(MAX_BYTES_BUILT_PER_EXECUTE - stored!.length);
    });

    it("Passes on a failure other than size while building a stream.", async () => {
        const context = makeContext({ messageRepo: { findOne: vi.fn().mockRejectedValue(new Error("database down")) } as any });
        await expect(openFastTransferHandle(context, 6, { type: "message", entityUid: "message:m1" }, {})).rejects.toThrow("database down");
        expect(context.session.handles[6]).toBeUndefined();
    });

    it("Parses a message body once per Execute however many streams open it, counting its bytes once.", async () => {
        const blobStore = new InMemoryBlobStore();
        await blobStore.put("bodies/m1", Buffer.from("Subject: Hi\r\n\r\nBody"));
        const findOne = vi.fn().mockResolvedValue({ uid: "m1", bodyBlobKey: "bodies/m1" });
        const budget = new ExecuteBudget();
        const context = makeContext({ blobStore, budget, messageRepo: { findOne } as any });

        const first = await loadStreamBody(context, 6, { type: "stream", entityUid: "message:m1", generation: "a" });
        const second = await loadStreamBody(context, 7, { type: "stream", entityUid: "message:m1", generation: "b" });

        expect(second).toEqual(first);
        expect(findOne).toHaveBeenCalledTimes(1);
        expect(budget.bytesRemaining).toBe(MAX_BYTES_BUILT_PER_EXECUTE - first.length);
    });
});

describe("Recipient lists", () => {
    it("Splits on semicolons only and parses addresses, display-name forms and bare names.", () => {
        expect(splitAddressList(undefined)).toEqual([]);
        expect(splitAddressList(" a@example.com ;; Doe, Jane <j@example.com> ")).toEqual(["a@example.com", "Doe, Jane <j@example.com>"]);
        expect(parseRecipientList('a@example.com; "Doe, Jane" <j@example.com>; Jane Doe; b@example.com, c@example.com')).toEqual([
            { name: "", address: "a@example.com" },
            { name: "Doe, Jane", address: "j@example.com" },
            { name: "Jane Doe" },
            { name: "", address: "b@example.com" },
            { name: "", address: "c@example.com" },
        ]);
        expect(parseAddressList("Jane <j@example.com>; Jane Doe; bad <x@>")).toEqual(["j@example.com"]);
        // A display name and a bare address separated by a comma within one entry belong together.
        expect(parseRecipientEntry("Jane Doe, j@example.com")).toEqual([{ name: "Jane Doe", address: "j@example.com" }]);
    });

    it("Marks unusable entries invalid and strips control characters from display names.", () => {
        expect(parseRecipientEntry("Jane <jane@>")).toEqual([{ name: "Jane", invalid: true }]);
        expect(parseRecipientEntry("<>")).toEqual([{ name: "<>", invalid: true }]);
        expect(parseRecipientEntry("to@example.com\r\nBcc: victim@example.com")[0].invalid).toBe(true);
        expect(parseRecipientEntry("Jane\tDoe <jane@example.com>")).toEqual([{ name: "Jane Doe", address: "jane@example.com" }]);
        expect(isPlainEmailAddress(`${"a".repeat(250)}@example.com`)).toBe(false);
    });

    it("Resolves recipients, reporting unresolved names and invalid entries separately.", async () => {
        const resolution = await resolveRecipientList("x@example.com; Nobody; <>", {
            mailboxUid: "mailbox-1",
            contactRepo: { find: vi.fn().mockResolvedValue([]) } as any,
        });
        expect(resolution).toEqual({ recipients: [{ name: "", address: "x@example.com" }], unresolved: ["Nobody"], invalid: ["<>"] });
    });
});
