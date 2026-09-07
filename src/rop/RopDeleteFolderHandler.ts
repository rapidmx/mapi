///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_DELETE_FOLDER = 0x1d;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused for an unrecognized `FolderId` - the same constant
 * `RopOpenFolderHandler` uses for its own analogous lookup miss. */
const ERROR_NOT_FOUND = 0x8004010f;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for both "the referenced handle isn't a folder"
 * and "the folder has content/subfolders but the matching cascade flag wasn't set" - the same constant reused
 * throughout this pragmatic subset as a generic "can't do this" signal, not a claim of exact per-condition
 * `[MS-OXCFOLD]` return-value-table parity (the real spec's own `ecFolderNotEmpty` isn't separately modeled). */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `DeleteFolderFlags` bits (`[MS-OXCFOLD]` §2.2.1.3.1, confirmed this session). */
const DEL_MESSAGES = 0x01;
const DEL_FOLDERS = 0x04;
const DELETE_HARD_DELETE = 0x10;

/**
 * `RopDeleteFolder` (`[MS-OXCFOLD]`/`[MS-OXCROPS]`, RopId `0x1D`): deletes a folder by `FolderId`. Reuses
 * `RecoverableRepoUtils.delete()` exactly as-built in Phase 2 for `Folder`/`Message`/`CalendarEvent` - see
 * `RopDeleteMessagesHandler.ts`'s own doc comment for why that matters (EAS's watermark-based incremental
 * delete detection depends on it).
 *
 * **Real spec semantics honored, not simplified away**: per `[MS-OXCFOLD]`, `RopDeleteFolder` only operates on
 * an empty folder by default - `DEL_MESSAGES` must be set to also delete the folder's own messages/calendar
 * events, and `DEL_FOLDERS` to also delete (recursively) its subfolders; a non-empty folder deleted without the
 * matching flag fails rather than silently cascading. `DELETE_HARD_DELETE` maps directly onto
 * `RepoDeleteOptions.purge` - real, exact semantic overlap with this library's own soft-delete model, not a
 * coincidence this codec papers over.
 *
 * The request's own `InputHandleIndex` (nominally the *parent* folder of the one being deleted) is validated to
 * be a real, open folder handle but not cross-checked against the target folder's actual `parentFolderUid` -
 * the same pragmatic simplification `RopOpenMessageHandler`'s own doc comment documents for its analogous
 * `FolderId` field.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopDeleteFolderHandler implements RopHandler {
    public readonly ropId = ROP_ID_DELETE_FOLDER;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const deleteFolderFlags: number = reader.readUInt8();
        const folderId: bigint = reader.readBigUInt64LE();

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "folder") {
            writer.writeUInt8(ROP_ID_DELETE_FOLDER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const target: string | undefined = context.session.folderIds[folderId.toString()];
        if (!target?.startsWith("folder:")) {
            writer.writeUInt8(ROP_ID_DELETE_FOLDER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_NOT_FOUND);
            return;
        }
        const uid = target.slice("folder:".length);
        const purge = (deleteFolderFlags & DELETE_HARD_DELETE) !== 0;

        const messages = await context.messageRepo.find({ folderUid: uid }, { ignoreACL: true });
        const events = await context.calendarEventRepo.find({ folderUid: uid }, { ignoreACL: true });
        const childFolders = await context.folderRepo.find({ parentFolderUid: uid }, { ignoreACL: true });

        if ((messages.length > 0 || events.length > 0) && (deleteFolderFlags & DEL_MESSAGES) === 0) {
            writer.writeUInt8(ROP_ID_DELETE_FOLDER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }
        if (childFolders.length > 0 && (deleteFolderFlags & DEL_FOLDERS) === 0) {
            writer.writeUInt8(ROP_ID_DELETE_FOLDER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        for (const message of messages) {
            await context.messageRepo.delete(message.uid, { ignoreACL: true, purge });
        }
        for (const event of events) {
            await context.calendarEventRepo.delete(event.uid, { ignoreACL: true, purge });
        }
        for (const childFolder of childFolders) {
            await this.deleteFolderRecursive(childFolder.uid, context, purge);
        }
        await context.folderRepo.delete(uid, { ignoreACL: true, purge });

        writer.writeUInt8(ROP_ID_DELETE_FOLDER);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(0); // PartialCompletion - always false; a partial failure above already returned early
    }

    /** Cascades a `DEL_FOLDERS` delete into `folderUid`'s own messages/calendar events/subfolders before
     * deleting it - this data model's folder hierarchy is a plain parent-pointer tree (never a graph), so
     * unbounded recursion here needs no cycle protection, the same assumption `FolderTarget.resolveFolderChildren`
     * already makes. */
    private async deleteFolderRecursive(folderUid: string, context: RopContext, purge: boolean): Promise<void> {
        const messages = await context.messageRepo.find({ folderUid }, { ignoreACL: true });
        const events = await context.calendarEventRepo.find({ folderUid }, { ignoreACL: true });
        const childFolders = await context.folderRepo.find({ parentFolderUid: folderUid }, { ignoreACL: true });

        for (const message of messages) {
            await context.messageRepo.delete(message.uid, { ignoreACL: true, purge });
        }
        for (const event of events) {
            await context.calendarEventRepo.delete(event.uid, { ignoreACL: true, purge });
        }
        for (const childFolder of childFolders) {
            await this.deleteFolderRecursive(childFolder.uid, context, purge);
        }
        await context.folderRepo.delete(folderUid, { ignoreACL: true, purge });
    }
}
