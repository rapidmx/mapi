///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { resolveMessageBodyBytes } from "./MessageBodyStream.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_READ_STREAM = 0x2c;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a stream (or
 * doesn't exist)" - the same constant `RopGetHierarchyTableHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `0xBABE`: `ByteCount`'s sentinel value signaling "the real limit is the following 4-byte `MaximumByteCount`
 * field instead" (`[MS-OXCPRPT]`'s own `RopReadStream` request-buffer page). */
const BYTE_COUNT_USE_MAXIMUM = 0xbabe;

/**
 * `RopReadStream` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): reads up to `ByteCount` (or `MaximumByteCount`, if
 * `ByteCount` is the `0xBABE` sentinel) bytes from an already-opened stream (`RopOpenStream`), advancing the
 * stream's read position. Unlike most other ROPs in this pragmatic subset, `[MS-OXCROPS]` documents only one
 * combined response-buffer shape for this ROP (no separate Success/Failure pages) - `DataSize`/`Data` are
 * always present, `DataSize = 0` (empty `Data`) standing in for the failure case rather than the fields being
 * omitted entirely.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopReadStreamHandler implements RopHandler {
    public readonly ropId = ROP_ID_READ_STREAM;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const byteCount: number = reader.readUInt16LE();
        const requestedCount: number = byteCount === BYTE_COUNT_USE_MAXIMUM ? reader.readUInt32LE() : byteCount;

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "stream") {
            writer.writeUInt8(ROP_ID_READ_STREAM);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            writer.writeUInt16LE(0); // DataSize - see class doc comment on why this is still written on failure
            return;
        }

        const bytes = await resolveMessageBodyBytes(handle.entityUid, context.messageRepo, context.blobStore);
        const position = handle.streamPosition ?? 0;
        const slice = bytes.subarray(position, position + requestedCount);
        handle.streamPosition = position + slice.length;

        writer.writeUInt8(ROP_ID_READ_STREAM);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(slice.length); // DataSize
        writer.writeBytes(slice);
    }
}
