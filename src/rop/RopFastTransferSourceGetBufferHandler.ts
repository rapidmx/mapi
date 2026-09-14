///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { loadFastTransferBuffer } from "./FastTransferStream.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_GET_BUFFER = 0x4e;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a
 * `fastTransfer` handle (or doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its
 * own analogous check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/** The `BufferSize` sentinel (`[MS-OXCROPS]`'s own documented value) meaning "let the server pick the chunk
 * size". */
const BUFFER_SIZE_SERVER_DETERMINED = 0xbabe;

/** The response's own `TransferBufferSize` field is a 16-bit count (`[MS-OXCFXICS]`'s own `RopFastTransferSourceGetBuffer`
 * response-buffer page), so a single call can never actually return more than this many bytes. For the
 * `0xBABE` "server-determined" sentinel this caps what would otherwise be "the entire remaining buffer in one
 * call" (only spec-valid when that buffer happens to fit in a `uint16`) down to the same paging behavior a
 * client-specified `BufferSize` already gets - `writeUInt16LE` throwing `RangeError` for an out-of-range value
 * is not spec-valid for any size of transfer. */
const MAX_TRANSFER_BUFFER_SIZE = 0xffff;

/** `TransferStatus` values (`[MS-OXCFXICS]` §2.2.3.1.1.5.2, confirmed this session) - only the two this
 * pragmatic subset (which never errors mid-transfer once a `"fastTransfer"` handle exists, and never returns
 * `NoRoom`) ever produces. */
const TRANSFER_STATUS_PARTIAL = 0x0001;
const TRANSFER_STATUS_DONE = 0x0003;

/**
 * `RopFastTransferSourceGetBuffer` (`[MS-OXCFXICS]`/`[MS-OXCROPS]`, RopId `0x4E`): pages the FastTransfer
 * stream a prior `RopFastTransferSourceCopyTo`/`CopyProperties` built (`FastTransferStream.ts`) out of its
 * `"fastTransfer"` handle, `BufferSize` bytes at a time (or the entire remaining buffer, capped to
 * `MAX_TRANSFER_BUFFER_SIZE`, for the `0xBABE` "server-determined" sentinel - `MaximumBufferSize`, present only
 * in that case, is decoded to advance the reader correctly but not honored, since honoring it would mean
 * returning *more* than one chunk can carry, not less). Reports `Done` once the whole buffer has been returned
 * across one or more calls, `Partial` otherwise - `NoRoom` and `Error` are never produced (a `"fastTransfer"`
 * handle, once created, always has a complete, already-valid buffer to page from).
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

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
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

        const fullBuffer = await loadFastTransferBuffer(context, inputHandleIndex, handle);
        const position = handle.transferPosition ?? 0;
        const remaining = fullBuffer.length - position;
        const requestedSize = bufferSize === BUFFER_SIZE_SERVER_DETERMINED ? remaining : bufferSize;
        const chunkSize = Math.min(requestedSize, remaining, MAX_TRANSFER_BUFFER_SIZE);
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
