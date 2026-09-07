///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_OPEN_FOLDER = 0x02;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused as this handler's only failure `ReturnValue` (an
 * unrecognized `FolderId` - one not present in `session.folderIds`, i.e. not something `RopLogon` handed out
 * this session). */
const ERROR_NOT_FOUND = 0x8004010f;

/**
 * `RopOpenFolder` (`[MS-OXCFOLD]`/`[MS-OXCROPS]`): opens a folder by FID, producing a new `"folder"` Server
 * object handle later ROPs (`RopGetHierarchyTable`, a future `RopGetContentsTable`) reference by
 * `InputHandleIndex`. The request's own `OpenModeFlags` (`OpenSoftDeleted`, ...) is decoded to advance past
 * it correctly but not acted on - this pragmatic subset has no soft-deleted-folder recovery flow.
 *
 * Private-mailbox-only response shape: `IsGhosted` is always `0` and its conditional
 * `ServerCount`/`CheapServerCount`/`Servers` fields are never present, since `IsGhosted` and its trailing
 * fields are documented as public-folder-only (`[MS-OXCFOLD]`'s own "RopOpenFolder ROP Response Buffer"
 * page). `HasRules` is always `0` - this library has no `[MS-OXORULE]` rules-engine support.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopOpenFolderHandler implements RopHandler {
    public readonly ropId = ROP_ID_OPEN_FOLDER;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        reader.readUInt8(); // InputHandleIndex - the logon handle this open is performed against; only one
        // logon exists per session in this pragmatic subset, so it's consumed to advance the reader correctly
        // but never separately validated.
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // OpenModeFlags - OpenSoftDeleted not supported in this pragmatic subset
        const folderId: bigint = reader.readBigUInt64LE();

        const target: string | undefined = context.session.folderIds[folderId.toString()];
        if (!target) {
            writer.writeUInt8(ROP_ID_OPEN_FOLDER);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_NOT_FOUND);
            return;
        }

        context.session.handles[outputHandleIndex] = { type: "folder", entityUid: target };

        writer.writeUInt8(ROP_ID_OPEN_FOLDER);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(0); // HasRules - no rules-engine support
        writer.writeUInt8(0); // IsGhosted - always false for a private mailbox
    }
}
