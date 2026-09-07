///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_RELEASE = 0x01;

/**
 * `RopRelease` (`[MS-OXCROPS]`): releases a Server object handle. Confirmed against the spec's own real
 * captured example (`08 00 01 00 00 01 00 01 6F 00 00 00 6E 00 00 00`) that the request is exactly 3 bytes
 * (`RopId`/`LogonId`/`InputHandleIndex`) and, uniquely among ROPs, **produces no response entry at all** - the
 * server MUST NOT write anything to the output buffer for it, so `handle()` never touches `writer`.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopReleaseHandler implements RopHandler {
    public readonly ropId = ROP_ID_RELEASE;

    public handle(reader: BufferReader, _writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track per-logon-id state separately
        const inputHandleIndex: number = reader.readUInt8();
        delete context.session.handles[inputHandleIndex];
    }
}
