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

/** `MAPI_E_TOO_BIG`, for a write that would take the stream past `MAX_WRITE_STREAM_BYTES`. */
const ERROR_TOO_BIG = 0x80040305;

/** The largest body a write stream accumulates. The buffer lives in session state, which is re-saved on every
 * `Execute`, so it has to stay small. */
export const MAX_WRITE_STREAM_BYTES = 4 * 1024 * 1024;

/**
 * Appends `data` to the base64 string `existing`, touching only its last 4-character quantum. Base64 of a byte
 * count divisible by 3 has no padding and concatenates cleanly; only a padded final quantum (1 or 2 bytes) has to
 * be decoded and re-encoded together with the new bytes. Re-encoding the whole accumulated buffer on every write
 * made a body written in many chunks quadratic.
 */
export function appendBase64(existing: string, data: Buffer): string {
    if (!existing.endsWith("=")) {
        return existing + data.toString("base64");
    }
    const tail = Buffer.from(existing.slice(-4), "base64");
    return existing.slice(0, -4) + Buffer.concat([tail, data]).toString("base64");
}

/**
 * `RopWriteStream` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): writes bytes to a stream opened in `ReadWrite`/`Create`
 * mode (`RopOpenStreamHandler`'s write-mode branch - the only kind of writable stream this pragmatic subset
 * ever produces, always a `RopCreateMessage` draft's `PidTagBody`). Accumulates `Data` into the stream handle's
 * `writeBufferBase64` field (see `appendBase64` for how that avoids re-encoding everything written so far), up
 * to `MAX_WRITE_STREAM_BYTES`; a write past that fails with `MAPI_E_TOO_BIG` and nothing is appended.
 * `RopSubmitMessageHandler` decodes this buffer back to the final body text once composing is complete.
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

        const existing: string = handle.writeBufferBase64 ?? "";
        const existingSize: number = handle.writeSize ?? Buffer.byteLength(existing, "base64");
        if (existingSize + data.length > MAX_WRITE_STREAM_BYTES) {
            writer.writeUInt8(ROP_ID_WRITE_STREAM);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_TOO_BIG);
            writer.writeUInt16LE(0); // WrittenSize
            return;
        }
        handle.writeBufferBase64 = appendBase64(existing, data);
        handle.writeSize = existingSize + data.length;

        writer.writeUInt8(ROP_ID_WRITE_STREAM);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(data.length); // WrittenSize
    }
}
