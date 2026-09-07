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

export function encodeRopBuffer(buf: RopBuffer): Buffer {
    const writer = new BufferWriter();
    // RopSize covers itself (2 bytes) plus ropsList - the handle table isn't part of RopSize's count.
    writer.writeUInt16LE(2 + buf.ropsList.length);
    writer.writeBytes(buf.ropsList);
    for (const handle of buf.handleTable) {
        writer.writeUInt32LE(handle);
    }
    return writer.toBuffer();
}

export function decodeRopBuffer(buffer: Buffer): RopBuffer {
    const reader = new BufferReader(buffer);
    const ropSize = reader.readUInt16LE();
    const ropsList = reader.readBytes(ropSize - 2);
    const handleTable: number[] = [];
    while (reader.hasMore()) {
        handleTable.push(reader.readUInt32LE());
    }
    return { ropsList, handleTable };
}
