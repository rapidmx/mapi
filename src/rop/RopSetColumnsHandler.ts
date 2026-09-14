///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { readPropertyTagArray } from "../codec/PropertyValue.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_SET_COLUMNS = 0x12;

/** `TBLSTAT_COMPLETE` - this pragmatic subset's tables are always populated synchronously, so `RopSetColumns`
 * (and every other table ROP) never reports any other `TableStatus` value. */
const TABLE_STATUS_COMPLETE = 0x00;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a table (or
 * doesn't exist)" - the same constant `RopQueryRowsHandler`/`RopGetHierarchyTableHandler` use for their own
 * analogous checks. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_TOO_BIG`, for more than `MAX_PROPERTY_TAG_COUNT` columns. */
const ERROR_TOO_BIG = 0x80040305;

/**
 * `RopSetColumns` (`[MS-OXCTABL]`/`[MS-OXCROPS]`): configures which properties a table's subsequent
 * `RopQueryRows` calls return, and in what order. Unlike `RopOpenFolder`/`RopGetHierarchyTable`, this ROP
 * doesn't create a new handle - it mutates the existing table handle referenced by `InputHandleIndex` in
 * place, so its response echoes `InputHandleIndex` rather than an `OutputHandleIndex`, the same convention
 * `RopQueryRows`' own response uses (confirmed against `[MS-OXCTABL]`'s real captured response example).
 *
 * Validates that `InputHandleIndex` actually refers to a `"table"` handle before mutating it, the same check
 * every other handle-consuming ROP in this pragmatic subset already makes - a missing or wrong-type handle
 * previously fell through silently (no `else` branch) to a bare success response, meaning a client that set
 * columns against a stale/released/wrong-kind handle index only discovered the real problem one ROP later,
 * when `RopQueryRows` on that same index returned `MAPI_E_INVALID_OBJECT` instead.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSetColumnsHandler implements RopHandler {
    public readonly ropId = ROP_ID_SET_COLUMNS;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // SetColumnsFlags - this pragmatic subset has no async/deferred column-set variant
        const columns = readPropertyTagArray(reader, reader.readUInt16LE());

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "table" || !columns) {
            writer.writeUInt8(ROP_ID_SET_COLUMNS);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(columns ? ERROR_INVALID_OBJECT : ERROR_TOO_BIG);
            return;
        }
        handle.columns = columns;

        writer.writeUInt8(ROP_ID_SET_COLUMNS);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(TABLE_STATUS_COMPLETE);
    }
}
