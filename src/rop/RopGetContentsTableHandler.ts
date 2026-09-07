///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { Folder, FolderType } from "@rapidmx/restapi";
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { resolveFolderCalendarEvents } from "./CalendarEventTarget.js";
import { resolveFolderMessages } from "./MessageTarget.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_CONTENTS_TABLE = 0x05;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a folder (or
 * doesn't exist)" - the same constant `RopGetHierarchyTableHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopGetContentsTable` (`[MS-OXCFOLD]`/`[MS-OXCROPS]`): opens a table listing an already-opened folder's
 * messages, producing a new `"table"` Server object handle `RopSetColumns`/`RopQueryRows` reference by
 * `InputHandleIndex` - the exact same request/response shape as `RopGetHierarchyTable` (confirmed identical
 * field-by-field against `[MS-OXCROPS]`'s own request-buffer page, differing only in `RopId`), just listing
 * messages instead of child folders. `RopSetColumns`/`RopQueryRows` themselves need no changes at all to
 * support this - they already operate generically on whatever `rows`/`columns` a table handle carries,
 * dispatching to `FolderTarget`'s or `MessageTarget`'s resolvers by row-target prefix (`"folder:"`/`"virtual:"`
 * vs `"message:"`).
 *
 * A folder's contents table is only ever non-empty for a real folder (`"folder:<uid>"`) - the virtual special
 * folders (`RopLogonHandler`'s own doc comment) have no real backing row to hold messages under, so opening a
 * contents table on one always yields an empty table rather than an error.
 *
 * **Calendar folders** (`Folder.type === FolderType.CALENDAR`) list `"calendarEvent:<uid>"` rows resolved via
 * `calendarEventRepo` instead of `"message:<uid>"` rows - a single table handle only ever holds one kind of
 * row, decided once here by checking the folder's own `type`, exactly the same "resolve by target-string
 * prefix downstream" design `MessageTarget`/`FolderTarget` already use for their own rows.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopGetContentsTableHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_CONTENTS_TABLE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // TableFlags - only the Standard (0x00) case is supported; see RopGetHierarchyTableHandler's own doc comment

        const folderHandle = context.session.handles[inputHandleIndex];
        if (!folderHandle || folderHandle.type !== "folder") {
            writer.writeUInt8(ROP_ID_GET_CONTENTS_TABLE);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const rows: string[] = folderHandle.entityUid.startsWith("folder:")
            ? await this.resolveRows(folderHandle.entityUid.slice("folder:".length), context)
            : [];
        context.session.handles[outputHandleIndex] = { type: "table", entityUid: folderHandle.entityUid, rows, cursor: 0 };

        writer.writeUInt8(ROP_ID_GET_CONTENTS_TABLE);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }

    private async resolveRows(folderUid: string, context: RopContext): Promise<string[]> {
        const folder: Folder | undefined = await context.folderRepo.findOne(folderUid, { ignoreACL: true });
        if (folder?.type === FolderType.CALENDAR) {
            return resolveFolderCalendarEvents(folderUid, context.calendarEventRepo);
        }
        return resolveFolderMessages(folderUid, context.messageRepo);
    }
}
