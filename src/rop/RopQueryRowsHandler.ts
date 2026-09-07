///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { writePropertyValue } from "../codec/PropertyValue.js";
import { resolvePropertyValues } from "./PropertyResolvers.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_QUERY_ROWS = 0x15;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a table (or
 * doesn't exist)" - the same constant `RopGetHierarchyTableHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `BOOKMARK_END` (`[MS-OXCTABL]`'s `Origin` field) - this pragmatic subset has no separate seek/bookmark ROP
 * support, so every response reports the cursor as having landed at the table's current end. */
const ORIGIN_BOOKMARK_END = 0x02;

/**
 * `RopQueryRows` (`[MS-OXCTABL]`/`[MS-OXCROPS]`): fetches up to `RowCount` rows from an already-configured
 * table (`RopGetHierarchyTable` + `RopSetColumns`), advancing the table's cursor. Confirmed field-by-field
 * against `[MS-OXCTABL]`'s own real captured request/response byte example - including the response's
 * `RowData` encoding (`PropertyRow`: a leading `Flags` byte, `0x00` for a `StandardPropertyRow` where every
 * column is a plain `PropertyValue`, `0x01` for a `FlaggedPropertyRow` where each column gets its own
 * `FlaggedPropertyValue` flag+optional-value). This pragmatic subset always emits `StandardPropertyRow`
 * (`Flags = 0x00`) - every configured column always resolves to *some* value here (falling back to a
 * type-appropriate zero/empty default for an unsupported property, never an error), so `FlaggedPropertyRow`'s
 * per-column error signaling is never needed.
 *
 * Backward reads (`ForwardRead = FALSE`) and the `NoAdvance` flag are decoded (to advance the reader
 * correctly) but not honored - this pragmatic subset's tables are simple forward-only cursors. Works generically
 * against either a `RopGetHierarchyTable` (folder rows, `"folder:"`/`"virtual:"` targets) or a
 * `RopGetContentsTable` (message rows, `"message:"` targets) handle - dispatching to `FolderTarget`'s or
 * `MessageTarget`'s resolver and property-value mapping by row-target prefix, since a single table handle only
 * ever holds one kind of row.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopQueryRowsHandler implements RopHandler {
    public readonly ropId = ROP_ID_QUERY_ROWS;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // QueryRowsFlags - NoAdvance/EnablePackedBuffers not supported; always advances
        reader.readUInt8(); // ForwardRead - backward reads not supported; always reads forward
        const requestedCount: number = reader.readUInt16LE();

        const table = context.session.handles[inputHandleIndex];
        if (!table || table.type !== "table") {
            writer.writeUInt8(ROP_ID_QUERY_ROWS);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const rows: string[] = table.rows ?? [];
        const cursor: number = table.cursor ?? 0;
        const slice: string[] = rows.slice(cursor, cursor + requestedCount);
        table.cursor = cursor + slice.length;

        const rowBuffers: Buffer[] = [];
        for (const target of slice) {
            rowBuffers.push(await this.buildRow(context, target, table.columns ?? []));
        }

        writer.writeUInt8(ROP_ID_QUERY_ROWS);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(ORIGIN_BOOKMARK_END);
        writer.writeUInt16LE(slice.length);
        for (const rowBuffer of rowBuffers) {
            writer.writeBytes(rowBuffer);
        }
    }

    private async buildRow(
        context: RopContext,
        target: string,
        columns: { propertyId: number; propertyType: number }[],
    ): Promise<Buffer> {
        const writer = new BufferWriter();
        writer.writeUInt8(0x00); // Flags - StandardPropertyRow, see class doc comment
        const values = await resolvePropertyValues(target, columns, context);
        columns.forEach((column, index) => writePropertyValue(writer, column.propertyType, values[index]));
        return writer.toBuffer();
    }
}
