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

/** The response's own `TransferBufferSize` field is a 16-bit count, so a single call never returns more than this. */
const MAX_TRANSFER_BUFFER_SIZE = 0xffff;

/** `MAPI_E_TOO_BIG`: rebuilding the stream passed `MAX_FAST_TRANSFER_BYTES`. */
const ERROR_TOO_BIG = 0x80040305;

/** `MAPI_E_CALL_FAILED`: the stream was lost part-way through paging and can't be continued - see
 * `loadFastTransferBuffer`. The client starts the transfer again. */
const ERROR_TRANSFER_LOST = 0x80004005;

/** The response bytes before `TransferBuffer`: RopId, InputHandleIndex, ReturnValue, TransferStatus,
 * InProgressCount, TotalStepCount, Reserved, TransferBufferSize. */
const RESPONSE_HEADER_BYTES = 15;

/** `TransferStatus` values (`[MS-OXCFXICS]` §2.2.3.1.1.5.2). */
const TRANSFER_STATUS_PARTIAL = 0x0001;
const TRANSFER_STATUS_NO_ROOM = 0x0002;
const TRANSFER_STATUS_DONE = 0x0003;

/**
 * `RopFastTransferSourceGetBuffer` (`[MS-OXCFXICS]`/`[MS-OXCROPS]`, RopId `0x4E`): pages the FastTransfer
 * stream a prior `RopFastTransferSourceCopyTo`/`CopyProperties` built (`FastTransferStream.ts`) out of its
 * `"fastTransfer"` handle, `BufferSize` bytes at a time. For the `0xBABE` "server-determined" sentinel the server
 * picks the size, up to the client's `MaximumBufferSize`.
 *
 * Every chunk is also held to the room left in this request's ROP output buffer (`context.ropOutputRemaining`) and
 * the 16-bit `TransferBufferSize`. Reports `Done` once the whole stream has been returned, `Partial` otherwise, and
 * `NoRoom` when the output buffer has no space for any bytes this time (the client asks again in a new request).
 *
 * Fails with `MAPI_E_TOO_BIG` when a stream that had to be rebuilt is too large, and with `MAPI_E_CALL_FAILED` when
 * the stream was lost part-way through (see `loadFastTransferBuffer`) - never with bytes from a different stream.
 * Failure responses carry only `ReturnValue`, like every other failing handler here. `BackoffTime` is never emitted.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopFastTransferSourceGetBufferHandler implements RopHandler {
    public readonly ropId = ROP_ID_GET_BUFFER;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const bufferSize: number = reader.readUInt16LE();
        const requestedSize = bufferSize === BUFFER_SIZE_SERVER_DETERMINED ? reader.readUInt16LE() : bufferSize;

        const handle = context.session.handles[inputHandleIndex];
        const loaded = handle?.type === "fastTransfer" ? await loadFastTransferBuffer(context, inputHandleIndex, handle) : undefined;
        if (!loaded || typeof loaded === "string") {
            writer.writeUInt8(ROP_ID_GET_BUFFER);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(!loaded ? ERROR_INVALID_OBJECT : loaded === "tooBig" ? ERROR_TOO_BIG : ERROR_TRANSFER_LOST);
            return;
        }

        const position = handle.transferPosition ?? 0;
        const remaining = loaded.length - position;
        const room = Math.max(0, (context.ropOutputRemaining ?? Infinity) - RESPONSE_HEADER_BYTES);
        const chunkSize = Math.min(requestedSize, remaining, MAX_TRANSFER_BUFFER_SIZE, room);
        const chunk = loaded.subarray(position, position + chunkSize);
        handle.transferPosition = position + chunk.length;
        const done = position + chunk.length >= loaded.length;
        const status = done ? TRANSFER_STATUS_DONE : room === 0 ? TRANSFER_STATUS_NO_ROOM : TRANSFER_STATUS_PARTIAL;

        writer.writeUInt8(ROP_ID_GET_BUFFER);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(status);
        writer.writeUInt16LE(0); // InProgressCount - no real progress tracking in this pragmatic subset
        writer.writeUInt16LE(1); // TotalStepCount - pragmatic constant, only ever used for progress-bar display
        writer.writeUInt8(0); // Reserved
        writer.writeUInt16LE(chunk.length); // TransferBufferSize
        writer.writeBytes(chunk); // TransferBuffer
    }
}
