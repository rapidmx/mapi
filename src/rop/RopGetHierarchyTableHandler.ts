///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { resolveFolderChildren } from "./FolderTarget.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_HIERARCHY_TABLE = 0x04;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a folder
 * (or doesn't exist)". */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopGetHierarchyTable` (`[MS-OXCFOLD]`/`[MS-OXCROPS]`): opens a table listing an already-opened folder's
 * direct child folders, producing a new `"table"` Server object handle `RopSetColumns`/`RopQueryRows`
 * reference by `InputHandleIndex`. Only `TableFlags = 0x00` ("Standard") is supported - this pragmatic subset
 * has no deferred/async table population, so the immediate-row-count-returning variants some `TableFlags`
 * values request are moot, and the response omits the row-count field those variants add (confirmed absent
 * for the Standard case against `[MS-OXCROPS]`'s own request-buffer page, which explicitly lists `TableFlags`
 * as this ROP's only request payload beyond the common handle-index header).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopGetHierarchyTableHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_HIERARCHY_TABLE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // TableFlags - only the Standard (0x00) case is supported; see class doc comment

        const folderHandle = context.session.handles[inputHandleIndex];
        if (!folderHandle || folderHandle.type !== "folder") {
            writer.writeUInt8(ROP_ID_GET_HIERARCHY_TABLE);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const rows: string[] = await resolveFolderChildren(context.mailboxUid, folderHandle.entityUid, context.folderRepo);
        context.session.handles[outputHandleIndex] = { type: "table", entityUid: folderHandle.entityUid, rows, cursor: 0 };

        writer.writeUInt8(ROP_ID_GET_HIERARCHY_TABLE);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }
}
