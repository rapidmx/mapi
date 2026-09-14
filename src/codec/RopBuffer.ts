///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";

/**
 * The outer `ROP input/output buffer` framing (`[MS-OXCROPS]` §2.2.1): `RopSize` (2 bytes, the size of
 * itself plus `RopsList`) + `RopsList` (variable) + `ServerObjectHandleTable` (the remaining bytes, each a
 * 32-bit Server object handle referenced by index from within the ROPs).
 *
 * Deliberately **ROP-agnostic**: `RopsList` is a concatenated sequence of individual ROP request/response
 * buffers with no per-ROP length prefix (unlike WBXML's tag-based self-description) - each ROP's own byte
 * layout is bespoke and keyed by its leading `RopId` byte, so splitting `RopsList` into individual ROPs
 * requires ROP-specific decode logic this generic framing codec can't provide. A `RopDispatcher` (a later
 * build step, once real `RopHandler`s exist) walks `ropsList` with a `BufferReader`, reading each ROP's
 * `RopId` and delegating to the matching handler, which itself knows how many bytes its own ROP consumes -
 * exactly the same "handler owns its own wire format" division of responsibility `EasCommandHandler` already
 * uses for EAS commands.
 *
 * @author Jean-Philippe Steinmetz
 */
export interface RopBuffer {
    /** Raw bytes of the concatenated ROP request/response entries (`RopsList`). */
    ropsList: Buffer;
    /** The `ServerObjectHandleTable` array - one 32-bit handle per referenced Server object, in the order
     * ROPs within `ropsList` reference them by index. */
    handleTable: number[];
}

/** The largest `ServerObjectHandleTable` accepted. Every ROP addresses a handle by a one-byte index, so a table
 * with more than 255 entries can't be referenced anyway; a larger one is only a way to make the server allocate
 * and echo back an oversized array. */
export const MAX_HANDLE_TABLE_ENTRIES = 255;

export function encodeRopBuffer(buf: RopBuffer): Buffer {
    const writer = new BufferWriter();
    // RopSize covers itself (2 bytes) plus ropsList - the handle table isn't part of RopSize's count.
    writer.writeUInt16LE(2 + buf.ropsList.length);
    writer.writeBytes(buf.ropsList);
    // One allocation for the whole table instead of one small buffer per entry.
    const table = Buffer.alloc(buf.handleTable.length * 4);
    buf.handleTable.forEach((handle, index) => table.writeUInt32LE(handle >>> 0, index * 4));
    writer.writeBytes(table);
    return writer.toBuffer();
}

/** Throws a `RangeError` for a malformed buffer: a `RopSize` smaller than its own field or larger than the
 * buffer, a handle table that isn't a whole number of 4-byte entries, or one with more than
 * `MAX_HANDLE_TABLE_ENTRIES` entries. */
export function decodeRopBuffer(buffer: Buffer): RopBuffer {
    const reader = new BufferReader(buffer);
    const ropSize = reader.readUInt16LE();
    const ropsList = reader.readBytes(ropSize - 2);
    const tableBytes = reader.remaining;
    if (tableBytes % 4 !== 0 || tableBytes / 4 > MAX_HANDLE_TABLE_ENTRIES) {
        throw new RangeError(`RopBuffer: invalid ServerObjectHandleTable of ${tableBytes} bytes.`);
    }
    const handleTable: number[] = [];
    while (reader.hasMore()) {
        handleTable.push(reader.readUInt32LE());
    }
    return { ropsList, handleTable };
}
