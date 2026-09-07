///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { readPropertyTag } from "../codec/PropertyValue.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_SET_COLUMNS = 0x12;

/** `TBLSTAT_COMPLETE` - this pragmatic subset's tables are always populated synchronously, so `RopSetColumns`
 * (and every other table ROP) never reports any other `TableStatus` value. */
const TABLE_STATUS_COMPLETE = 0x00;

/**
 * `RopSetColumns` (`[MS-OXCTABL]`/`[MS-OXCROPS]`): configures which properties a table's subsequent
 * `RopQueryRows` calls return, and in what order. Unlike `RopOpenFolder`/`RopGetHierarchyTable`, this ROP
 * doesn't create a new handle - it mutates the existing table handle referenced by `InputHandleIndex` in
 * place, so its response echoes `InputHandleIndex` rather than an `OutputHandleIndex`, the same convention
 * `RopQueryRows`' own response uses (confirmed against `[MS-OXCTABL]`'s real captured response example).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSetColumnsHandler implements RopHandler {
    public readonly ropId = ROP_ID_SET_COLUMNS;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // SetColumnsFlags - this pragmatic subset has no async/deferred column-set variant
        const propertyTagCount: number = reader.readUInt16LE();
        const columns: { propertyId: number; propertyType: number }[] = [];
        for (let i = 0; i < propertyTagCount; i++) {
            columns.push(readPropertyTag(reader));
        }

        const handle = context.session.handles[inputHandleIndex];
        if (handle) {
            handle.columns = columns;
        }

        writer.writeUInt8(ROP_ID_SET_COLUMNS);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(TABLE_STATUS_COMPLETE);
    }
}
