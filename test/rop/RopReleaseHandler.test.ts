///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// RopReleaseHandler is pure logic with no DI/DB dependency, tested directly here - same precedent as the
// other test/mapi/**/*.test.ts files.
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { RopReleaseHandler } from "../../src/rop/RopReleaseHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

describe("RopReleaseHandler Tests", () => {
    const makeContext = function (): RopContext {
        return {
            mailboxUid: "mailbox-1",
            userUid: "user-1",
            session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
            folderRepo: {} as any,
            messageRepo: {} as any,
            mailboxRepo: {} as any,
            folderClass: {} as any,
            messageClass: {} as any,
            scanPipeline: {} as any,
            mailTransport: {} as any,
            blobStore: {} as any,
        };
    };

    it("Matches the spec's own real captured example (two RopRelease requests, no response bytes).", () => {
        // From [MS-OXCROPS]'s "RopRelease ROP Request" scenario: `08 00 01 00 00 01 00 01 6F 00 00 00 6E 00 00 00`.
        const ropsList = Buffer.from("0100000100016F0000006E000000", "hex").subarray(0, 6);
        const handler = new RopReleaseHandler();
        expect(handler.ropId).toBe(0x01);

        const reader = new BufferReader(ropsList);
        const writer = new BufferWriter();
        const context = makeContext();
        context.session.handles[0] = { type: "logon", entityUid: "mailbox-1" };
        context.session.handles[1] = { type: "folder", entityUid: "folder-1" };

        reader.readUInt8(); // RopId, consumed by the dispatcher in real use
        handler.handle(reader, writer, context);
        expect(context.session.handles[0]).toBeUndefined();
        expect(context.session.handles[1]).toEqual({ type: "folder", entityUid: "folder-1" });

        reader.readUInt8(); // RopId of the second RopRelease
        handler.handle(reader, writer, context);
        expect(context.session.handles[1]).toBeUndefined();

        // Uniquely among ROPs, RopRelease produces no response bytes at all.
        expect(writer.toBuffer().length).toBe(0);
        expect(reader.hasMore()).toBe(false);
    });

    it("Is a no-op when releasing a handle index that was never assigned.", () => {
        const handler = new RopReleaseHandler();
        const reader = new BufferReader(Buffer.from([0x00, 0x05]));
        const writer = new BufferWriter();
        const context = makeContext();

        expect(() => handler.handle(reader, writer, context)).not.toThrow();
        expect(writer.toBuffer().length).toBe(0);
    });
});
