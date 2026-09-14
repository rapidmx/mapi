///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { encodeGuid } from "../codec/MapiGuid.js";
import { Folder, FolderType } from "@rapidmx/restapi";
import { assignHandle } from "../MapiSessionManager.js";
import { assignOrGetFid } from "./FolderTarget.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_LOGON = 0xfe;

/** A fixed replica identity shared by every logon in this deployment - legitimate per `[MS-OXCSTOR]`'s own
 * "RopLogon ROP Success Response Buffer for Private Mailbox" section ("If the client did not set
 * USE_PER_MDB_REPLID_MAPPING... this value MUST be identical for all private mailbox logons on the same RPC
 * session"); this pragmatic subset never varies replica identity, so one fixed constant satisfies that. */
const REPL_ID = 1;
const REPL_GUID = "00000000-0000-0000-0000-000000000001";

/**
 * Order matches `[MS-OXCSTOR]` §2.2.1.1.3 ("RopLogon ROP Success Response Buffer for Private Mailbox")'s
 * `FolderIds` list exactly: Root, Deferred Action, Spooler Queue, IPM Subtree, Inbox, Outbox, Sent Items,
 * Deleted Items, Common Views, Schedule, Search, Views, Shortcuts (13 folders total, confirmed against the
 * spec page directly). Folders with a matching `FolderType` resolve to a real `Folder` when one exists for
 * this mailbox; the rest are virtual placeholders. Root/IPM Subtree have no real backing row at all in this
 * data model, matching how EAS's own `FolderSync` already represents "the root" as `ParentId "0"` rather than
 * a real `Folder` - not a new gap this ROP introduces. Deferred Action/Spooler Queue/Common Views/Schedule/
 * Search/Views/Shortcuts are MAPI-internal administrative folders this library's `FolderType` enum has no
 * concept of at all (it models the REST/EAS-relevant folder kinds - Inbox/Sent/Drafts/Deleted/Outbox/Junk/
 * Calendar/Contacts/Tasks/Notes/User - not Exchange's own internal bookkeeping folders). A real client rarely
 * if ever tries to open most of these directly; documented gap, not silently dropped.
 */
const SPECIAL_FOLDERS: { name: string; folderType?: FolderType }[] = [
    { name: "root" },
    { name: "deferredAction" },
    { name: "spoolerQueue" },
    { name: "ipmSubtree" },
    { name: "inbox", folderType: FolderType.INBOX },
    { name: "outbox", folderType: FolderType.OUTBOX },
    { name: "sentItems", folderType: FolderType.SENT_ITEMS },
    { name: "deletedItems", folderType: FolderType.DELETED_ITEMS },
    { name: "commonViews" },
    { name: "schedule" },
    { name: "search" },
    { name: "views" },
    { name: "shortcuts" },
];

/**
 * `RopLogon` (`[MS-OXCSTOR]`/`[MS-OXCROPS]`): the first ROP of any session, a prerequisite for every other
 * ROP. This pragmatic subset always performs a private-mailbox logon against the mailbox `BaseMapiEmsmdbRoute`
 * already resolved and authorized at `Connect` time - the request's own `Essdn` field is decoded (to advance
 * past it correctly) but never trusted for mailbox identity, the same principle `BaseMapiEmsmdbRoute.Connect`
 * already applies to that request's `UserDn` field. Because mailbox resolution already happened upstream,
 * this handler has no real failure path to model (`ecUnknownUser`/`ecLoginFailure`/... are all deferred).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopLogonHandler implements RopHandler {
    public readonly ropId = ROP_ID_LOGON;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const outputHandleIndex: number = reader.readUInt8();
        const logonFlags: number = reader.readUInt8();
        reader.readUInt32LE(); // OpenFlags - unused; every logon behaves as a plain private-mailbox logon
        reader.readUInt32LE(); // StoreState - per spec, unused and always 0
        const essdnSize: number = reader.readUInt16LE();
        if (essdnSize > 0) {
            reader.readBytes(essdnSize); // Essdn - never trusted for mailbox identity, see class doc comment
        }

        const fids: Record<string, number> = await this.assignFolderIds(context);
        assignHandle(context.session, outputHandleIndex, { type: "logon", entityUid: context.mailboxUid });

        writer.writeUInt8(ROP_ID_LOGON);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(logonFlags); // echoed back unchanged, per spec
        for (const folder of SPECIAL_FOLDERS) {
            writer.writeBigUInt64LE(BigInt(fids[folder.name]));
        }
        writer.writeUInt8(0x01); // ResponseFlags - Reserved bit only; no OwnerRight/SendAsRight/OOF claims in this subset
        writer.writeBytes(encodeGuid(context.mailboxUid));
        writer.writeUInt16LE(REPL_ID);
        writer.writeBytes(encodeGuid(REPL_GUID));
        writer.writeBytes(encodeLogonTime(new Date()));
        writer.writeBigUInt64LE(1n); // GwartTime - fixed constant; no real GWART/AD concept in this deployment
        writer.writeUInt32LE(0); // StoreState - unused
    }

    /**
     * Assigns this session's FIDs for the 13 special folders, remembering each one's real-`Folder`-or-virtual-
     * placeholder target in `session.folderIds` so a later `RopOpenFolder` can resolve it back to something
     * real. Delegates to `FolderTarget.assignOrGetFid` - the same "reuse an existing FID for this target if one
     * was already assigned, otherwise mint the next free one" logic `RopGetHierarchyTable`/`RopQueryRows` use
     * for child folders - rather than wholesale-replacing `session.folderIds` with a freshly-numbered map
     * starting back at FID 1. A second `RopLogon` on a live session (spec-legal - a real client can reconnect
     * or re-logon within one session context) previously invalidated every FID a client had already learned
     * for a child folder via `RopGetHierarchyTable`/`RopQueryRows` (all assigned starting at FID 14, since the
     * old code always reset the counter to 1-13 for just these specials) and could even cause FIDs 14+ to be
     * silently re-issued to a *different* folder than the one the client last associated with that number.
     * Reusing `assignOrGetFid` means a re-`RopLogon` is now purely additive: these 13 targets get their
     * existing FIDs back if this is a second logon, and any child-folder FIDs already assigned are left alone.
     */
    private async assignFolderIds(context: RopContext): Promise<Record<string, number>> {
        // The four real-folder-type lookups (Inbox/Outbox/Sent Items/Deleted Items) are independent of each
        // other, so resolve them concurrently rather than one sequential round trip per special folder -
        // FID *assignment* itself still happens afterward, in SPECIAL_FOLDERS' own fixed order, so which FID
        // ends up bound to which target is unaffected by fetch order.
        const targets = await Promise.all(
            SPECIAL_FOLDERS.map(async (folder): Promise<string> => {
                if (!folder.folderType) {
                    return `virtual:${folder.name}`;
                }
                // The oldest folder of the type, like restapi's findOrCreateWellKnownFolder, so a mailbox that somehow
                // has two gets the same one every time.
                const existing: Folder[] = await context.folderRepo.find(
                    { mailboxUid: context.mailboxUid, type: folder.folderType, sort: { dateCreated: "ASC", uid: "ASC" }, limit: 1 } as any,
                    { ignoreACL: true, limit: 1 },
                );
                return existing[0] ? `folder:${existing[0].uid}` : `virtual:${folder.name}`;
            }),
        );

        const fids: Record<string, number> = {};
        SPECIAL_FOLDERS.forEach((folder, index) => {
            fids[folder.name] = assignOrGetFid(context.session, targets[index]);
        });
        return fids;
    }
}

/** `LogonTime` structure (`[MS-OXCROPS]` "LogonTime Structure"): `Seconds`/`Minutes`/`Hour`/`DayOfWeek`
 * (1 byte each, `DayOfWeek` Sunday=0)/`Day`/`Month` (1 byte each, `Month` January=1)/`Year` (2 bytes) - all
 * UTC, confirmed against the spec's own byte-layout table. */
function encodeLogonTime(date: Date): Buffer {
    const writer = new BufferWriter();
    writer.writeUInt8(date.getUTCSeconds());
    writer.writeUInt8(date.getUTCMinutes());
    writer.writeUInt8(date.getUTCHours());
    writer.writeUInt8(date.getUTCDay());
    writer.writeUInt8(date.getUTCDate());
    writer.writeUInt8(date.getUTCMonth() + 1);
    writer.writeUInt16LE(date.getUTCFullYear());
    return writer.toBuffer();
}
