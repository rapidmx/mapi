///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./rop/RopHandler.js";

/**
 * Walks a decoded `RopBuffer`'s `ropsList` blob, dispatching each contained ROP - identified by its own
 * leading `RopId` byte - to the matching registered handler in turn, collecting every response into one
 * output buffer. Deliberately ROP-agnostic: it knows nothing about any specific ROP's own fields, only how to
 * find the *next* one, via the handler's own advanced `reader` position once `handle()` returns - see
 * `RopBuffer.ts`'s doc comment for why a fully generic split isn't possible for this wire format.
 *
 * An unrecognized `RopId` stops processing immediately: the bytes belonging to that ROP (and everything after
 * it) can't be skipped without knowing that ROP's own layout, so there is no safe way to resync and continue.
 * This is an inherent constraint of the wire format, not a design choice this library is choosing not to
 * handle - a real client only ever sends ROPs the server has already told it (via its supported-command
 * surface) that it understands.
 *
 * @author Jean-Philippe Steinmetz
 */
export async function dispatchRops(
    ropsList: Buffer,
    handlers: Map<number, RopHandler>,
    context: RopContext,
): Promise<Buffer> {
    const reader = new BufferReader(ropsList);
    const writer = new BufferWriter();
    while (reader.hasMore()) {
        const ropId = reader.readUInt8();
        const handler = handlers.get(ropId);
        if (!handler) {
            break;
        }
        await handler.handle(reader, writer, context);
    }
    return writer.toBuffer();
}
