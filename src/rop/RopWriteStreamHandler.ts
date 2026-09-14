///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { MapiObjectHandle } from "../MapiSessionManager.js";
import { handleDataKey } from "./HandleDataCache.js";
import { handleDataStoreOf, type RopContext, type RopHandler } from "./RopHandler.js";

const ROP_ID_WRITE_STREAM = 0x2d;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a write-mode
 * stream (or doesn't exist)" - the same constant `RopReadStreamHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** `MAPI_E_TOO_BIG`, for a write that would take the stream past `MAX_WRITE_STREAM_BYTES`. */
const ERROR_TOO_BIG = 0x80040305;

/** The largest body a write stream accumulates. */
export const MAX_WRITE_STREAM_BYTES = 4 * 1024 * 1024;

/** How long a write stream's chunks are kept after its last write - the session's own idle lifetime. */
export const WRITE_STREAM_TTL_SECONDS = 15 * 60;

/** The `HandleDataStore` key of the write stream at `handleIndex`. */
export function writeStreamKey(sessionUid: string, handleIndex: number, generation: string | undefined): string {
    return `${handleDataKey(sessionUid, handleIndex, generation)}:write`;
}

/**
 * `RopWriteStream` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): writes bytes to a stream opened in `ReadWrite`/`Create`
 * mode (`RopOpenStreamHandler`'s write-mode branch - the only kind of writable stream this pragmatic subset
 * ever produces, always a `RopCreateMessage` draft's `PidTagBody`). Each write is stored as a chunk keyed by its
 * offset in the `HandleDataStore` (Redis when configured), not in the session: only `writeSize` is session state.
 * Keying by offset makes a retried write (after a request whose session save lost a conflict) replace its chunk
 * instead of appending it twice. Up to `MAX_WRITE_STREAM_BYTES`; a write past that fails with `MAPI_E_TOO_BIG` and
 * nothing is stored. `RopSubmitMessageHandler` reassembles the body (`readWriteStream`) once composing is complete.
 *
 * Like `RopReadStream`, `[MS-OXCROPS]` documents only one combined response-buffer shape for this ROP (no
 * separate Success/Failure pages) - `WrittenSize` is always present, `0` standing in for the failure case.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopWriteStreamHandler implements RopHandler {
    public readonly ropId = ROP_ID_WRITE_STREAM;
    public readonly failureTailBytes = 2; // DataSize/WrittenSize

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const dataSize: number = reader.readUInt16LE();
        const data: Buffer = reader.readBytes(dataSize);

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "stream" || handle.writeTargetHandleIndex === undefined) {
            writer.writeUInt8(ROP_ID_WRITE_STREAM);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            writer.writeUInt16LE(0); // WrittenSize - see class doc comment on why this is still written on failure
            return;
        }

        const existingSize: number = handle.writeSize ?? 0;
        if (existingSize + data.length > MAX_WRITE_STREAM_BYTES) {
            writer.writeUInt8(ROP_ID_WRITE_STREAM);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_TOO_BIG);
            writer.writeUInt16LE(0); // WrittenSize
            return;
        }
        if (data.length > 0) {
            await handleDataStoreOf(context).putChunk(
                writeStreamKey(context.session.uid, inputHandleIndex, handle.generation),
                existingSize,
                data,
                WRITE_STREAM_TTL_SECONDS,
            );
        }
        handle.writeSize = existingSize + data.length;

        writer.writeUInt8(ROP_ID_WRITE_STREAM);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(data.length); // WrittenSize
    }
}

/** The bytes written so far to the write stream `stream` at `handleIndex`, or `undefined` when they can no longer be
 * reassembled (a chunk expired or was evicted). */
export async function readWriteStream(context: RopContext, handleIndex: number, stream: MapiObjectHandle): Promise<Buffer | undefined> {
    return handleDataStoreOf(context).readChunks(writeStreamKey(context.session.uid, handleIndex, stream.generation), stream.writeSize ?? 0);
}
