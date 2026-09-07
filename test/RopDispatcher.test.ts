///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { dispatchRops } from "../src/RopDispatcher.js";
import { RopReleaseHandler } from "../src/rop/RopReleaseHandler.js";
import type { RopContext } from "../src/rop/RopHandler.js";
import { MapiSessionContext } from "../src/MapiSessionManager.js";

describe("RopDispatcher Tests", () => {
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

    it("Dispatches the spec's own real captured RopRelease example (two ROPs) to the registered handler.", async () => {
        // From [MS-OXCROPS]'s "RopRelease ROP Request" scenario: RopsList = `01 00 00 01 00 01` (RopSize/
        // ServerObjectHandleTable are RopBuffer.ts's concern, not the dispatcher's - see that file's tests).
        const ropsList = Buffer.from("010000010001", "hex");
        const handlers = new Map([[0x01, new RopReleaseHandler()]]);
        const context = makeContext();
        context.session.handles[0] = { type: "logon", entityUid: "mailbox-1" };
        context.session.handles[1] = { type: "folder", entityUid: "folder-1" };

        const response = await dispatchRops(ropsList, handlers, context);

        expect(response.length).toBe(0); // RopRelease produces no response bytes
        expect(context.session.handles[0]).toBeUndefined();
        expect(context.session.handles[1]).toBeUndefined();
    });

    it("Stops processing at the first unrecognized RopId, since its byte layout can't be skipped safely.", async () => {
        const ropsList = Buffer.from([0x99, 0xaa, 0xbb, 0x01, 0x00, 0x00]); // unknown RopId 0x99, then a real RopRelease
        const handlers = new Map([[0x01, new RopReleaseHandler()]]);
        const context = makeContext();
        context.session.handles[0] = { type: "logon", entityUid: "mailbox-1" };

        await dispatchRops(ropsList, handlers, context);

        // The trailing RopRelease is never reached - the unknown ROP's bytes couldn't be skipped past.
        expect(context.session.handles[0]).toEqual({ type: "logon", entityUid: "mailbox-1" });
    });

    it("Returns an empty buffer for an empty ropsList.", async () => {
        const response = await dispatchRops(Buffer.alloc(0), new Map(), makeContext());
        expect(response.length).toBe(0);
    });
});
