///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_WRITE_STREAM = 0x2d;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a write-mode
 * stream (or doesn't exist)" - the same constant `RopReadStreamHandler` uses for its own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopWriteStream` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): writes bytes to a stream opened in `ReadWrite`/`Create`
 * mode (`RopOpenStreamHandler`'s write-mode branch - the only kind of writable stream this pragmatic subset
 * ever produces, always a `RopCreateMessage` draft's `PidTagBody`). Accumulates `Data` into the stream handle's
 * `writeBufferBase64` field, decode-concat-reencode each call rather than a naive string concatenation of
 * per-call base64 chunks - base64 is not concatenation-safe (a chunk's padding only belongs at the true end of
 * the data), so reassembling the real bytes requires decoding what's accumulated so far, appending the new raw
 * bytes, and re-encoding the whole thing. `RopSaveChangesMessageHandler`/`RopSubmitMessageHandler` decode this
 * buffer back to the final body text once composing is complete.
 *
 * Like `RopReadStream`, `[MS-OXCROPS]` documents only one combined response-buffer shape for this ROP (no
 * separate Success/Failure pages) - `WrittenSize` is always present, `0` standing in for the failure case.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopWriteStreamHandler implements RopHandler {
    public readonly ropId = ROP_ID_WRITE_STREAM;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
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

        const existing: Buffer = handle.writeBufferBase64 ? Buffer.from(handle.writeBufferBase64, "base64") : Buffer.alloc(0);
        handle.writeBufferBase64 = Buffer.concat([existing, data]).toString("base64");

        writer.writeUInt8(ROP_ID_WRITE_STREAM);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(data.length); // WrittenSize
    }
}
