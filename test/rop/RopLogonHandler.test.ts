///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// RopLogonHandler is tested directly here against a fake folderRepo (a genuinely un-reachable real DB isn't
// needed for this pure request/response encoding logic) - real HTTP+DB round trips are covered separately in
// test/routes/mongo/MapiEmsmdbRoute.test.ts (and its sql/ counterpart).
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { decodeGuid } from "../../src/codec/MapiGuid.js";
import { RopLogonHandler } from "../../src/rop/RopLogonHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { FolderType } from "@rapidmx/restapi";

const MAILBOX_UID = "11111111-1111-1111-1111-111111111111";

function buildLogonRequest({ logonId = 0, outputHandleIndex = 0, logonFlags = 0x01, essdn = "" } = {}): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(outputHandleIndex);
    writer.writeUInt8(logonFlags);
    writer.writeUInt32LE(0); // OpenFlags
    writer.writeUInt32LE(0); // StoreState
    if (essdn.length > 0) {
        const essdnBytes = Buffer.concat([Buffer.from(essdn, "ascii"), Buffer.from([0x00])]);
        writer.writeUInt16LE(essdnBytes.length);
        writer.writeBytes(essdnBytes);
    } else {
        writer.writeUInt16LE(0);
    }
    return writer.toBuffer();
}

function makeContext(folderRepo: any): RopContext {
    return {
        mailboxUid: MAILBOX_UID,
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: MAILBOX_UID, userUid: "user-1" }),
        folderRepo,
        messageRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopLogonHandler Tests", () => {
    it("Has RopId 0xFE.", () => {
        expect(new RopLogonHandler().ropId).toBe(0xfe);
    });

    it("Produces a well-formed success response with 13 FIDs, all virtual when no real folders exist.", async () => {
        const folderRepo = { find: vi.fn().mockResolvedValue([]) };
        const context = makeContext(folderRepo);
        const handler = new RopLogonHandler();
        const reader = new BufferReader(buildLogonRequest({ outputHandleIndex: 3, logonFlags: 0x01 }));
        const writer = new BufferWriter();

        await handler.handle(reader, writer, context);
        const response = new BufferReader(writer.toBuffer());

        expect(response.readUInt8()).toBe(0xfe); // RopId
        expect(response.readUInt8()).toBe(3); // OutputHandleIndex
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt8()).toBe(0x01); // LogonFlags, echoed

        const fids: bigint[] = [];
        for (let i = 0; i < 13; i++) {
            fids.push(response.readBigUInt64LE());
        }
        expect(fids).toEqual([1n, 2n, 3n, 4n, 5n, 6n, 7n, 8n, 9n, 10n, 11n, 12n, 13n]);

        expect(response.readUInt8()).toBe(0x01); // ResponseFlags
        expect(decodeGuid(response)).toBe(MAILBOX_UID); // MailboxGuid
        expect(response.readUInt16LE()).toBe(1); // ReplId
        decodeGuid(response); // ReplGuid
        response.readBytes(8); // LogonTime
        expect(response.readBigUInt64LE()).toBe(1n); // GwartTime
        expect(response.readUInt32LE()).toBe(0); // StoreState
        expect(response.hasMore()).toBe(false);

        // All 13 folders are virtual placeholders since folderRepo.find() always returned [].
        expect(Object.values(context.session.folderIds).every((v) => v.startsWith("virtual:"))).toBe(true);
        expect(context.session.handles[3]).toEqual({ type: "logon", entityUid: MAILBOX_UID });
    });

    it("Resolves Inbox/Outbox/Sent Items/Deleted Items to a real Folder when one exists, leaving the rest virtual.", async () => {
        const byType: Record<string, string> = {
            [FolderType.INBOX]: "inbox-folder-uid",
            [FolderType.OUTBOX]: "outbox-folder-uid",
            [FolderType.SENT_ITEMS]: "sent-folder-uid",
            [FolderType.DELETED_ITEMS]: "deleted-folder-uid",
        };
        const folderRepo = {
            find: vi.fn().mockImplementation(async (query: any) => {
                const uid = byType[query.type];
                return uid ? [{ uid }] : [];
            }),
        };
        const context = makeContext(folderRepo);
        const handler = new RopLogonHandler();
        await handler.handle(new BufferReader(buildLogonRequest()), new BufferWriter(), context);

        const targets = Object.values(context.session.folderIds);
        expect(targets).toContain("folder:inbox-folder-uid");
        expect(targets).toContain("folder:outbox-folder-uid");
        expect(targets).toContain("folder:sent-folder-uid");
        expect(targets).toContain("folder:deleted-folder-uid");
        // Root remains virtual regardless - this data model has no real backing row for it.
        const rootFid = Object.keys(context.session.folderIds)[0];
        expect(context.session.folderIds[rootFid]).toBe("virtual:root");
    });

    it("Never trusts the request's Essdn field for mailbox identity - consumes it but ignores its value.", async () => {
        const folderRepo = { find: vi.fn().mockResolvedValue([]) };
        const context = makeContext(folderRepo);
        const handler = new RopLogonHandler();
        const reader = new BufferReader(buildLogonRequest({ essdn: "/o=SomeOtherOrg/cn=someone-else" }));
        const writer = new BufferWriter();

        await handler.handle(reader, writer, context);

        expect(reader.hasMore()).toBe(false);
        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue - still succeeds, using the session's own mailboxUid
    });
});
