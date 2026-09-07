///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_BUFFER = 0x4e;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a
 * `fastTransfer` handle (or doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its
 * own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** The `BufferSize` sentinel (`[MS-OXCROPS]`'s own documented value) meaning "let the server pick the chunk
 * size" - honored here as "return the whole remaining buffer in one call", the simplest spec-valid choice for
 * this pragmatic subset's already-fully-built-in-memory transfer buffers. */
const BUFFER_SIZE_SERVER_DETERMINED = 0xbabe;

/** `TransferStatus` values (`[MS-OXCFXICS]` §2.2.3.1.1.5.2, confirmed this session) - only the two this
 * pragmatic subset (which never errors mid-transfer once a `"fastTransfer"` handle exists, and never returns
 * `NoRoom`) ever produces. */
const TRANSFER_STATUS_PARTIAL = 0x0001;
const TRANSFER_STATUS_DONE = 0x0003;

/**
 * `RopFastTransferSourceGetBuffer` (`[MS-OXCFXICS]`/`[MS-OXCROPS]`, RopId `0x4E`): pages the FastTransfer
 * stream a prior `RopFastTransferSourceCopyTo`/`CopyProperties` built (`FastTransferStream.ts`) out of its
 * `"fastTransfer"` handle, `BufferSize` bytes at a time (or the entire remaining buffer at once for the
 * `0xBABE` "server-determined" sentinel - `MaximumBufferSize`, present only in that case, is decoded to advance
 * the reader correctly but not honored, since this pragmatic subset's transfer buffers are already small,
 * fully-in-memory test-mailbox-scale content, not something worth capping). Reports `Done` once the whole
 * buffer has been returned across one or more calls, `Partial` otherwise - `NoRoom` and `Error` are never
 * produced (a `"fastTransfer"` handle, once created, always has a complete, already-valid buffer to page from).
 *
 * `BackoffTime` is never emitted (this pragmatic subset never returns the one `ReturnValue` that field is
 * conditional on), and the failure path (`ERROR_INVALID_OBJECT`) omits every field after `ReturnValue` entirely
 * - the same "just the fixed header, no success-only tail" shape every other failing handler in this pragmatic
 * subset already uses, not a specific claim about the real spec's own error-path byte layout.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopFastTransferSourceGetBufferHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_BUFFER;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const bufferSize: number = reader.readUInt16LE();
        if (bufferSize === BUFFER_SIZE_SERVER_DETERMINED) {
            reader.readUInt16LE(); // MaximumBufferSize - see class doc comment
        }

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "fastTransfer") {
            writer.writeUInt8(ROP_ID_GET_BUFFER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const fullBuffer = Buffer.from(handle.transferBufferBase64 ?? "", "base64");
        const position = handle.transferPosition ?? 0;
        const remaining = fullBuffer.length - position;
        const chunkSize = bufferSize === BUFFER_SIZE_SERVER_DETERMINED ? remaining : Math.min(bufferSize, remaining);
        const chunk = fullBuffer.subarray(position, position + chunkSize);
        handle.transferPosition = position + chunk.length;
        const done = handle.transferPosition >= fullBuffer.length;

        writer.writeUInt8(ROP_ID_GET_BUFFER);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(done ? TRANSFER_STATUS_DONE : TRANSFER_STATUS_PARTIAL);
        writer.writeUInt16LE(0); // InProgressCount - no real progress tracking in this pragmatic subset
        writer.writeUInt16LE(1); // TotalStepCount - pragmatic constant, only ever used for progress-bar display
        writer.writeUInt8(0); // Reserved
        writer.writeUInt16LE(chunk.length); // TransferBufferSize
        writer.writeBytes(chunk); // TransferBuffer
    }
}
