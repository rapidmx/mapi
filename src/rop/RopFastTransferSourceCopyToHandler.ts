///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { readPropertyTag } from "../codec/PropertyValue.js";
import { buildFastTransferStream } from "./FastTransferStream.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_COPY_TO = 0x4d;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a folder or
 * message (or doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its own analogous
 * check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopFastTransferSourceCopyTo` (`[MS-OXCFXICS]`/`[MS-OXCROPS]`, RopId `0x4D`): begins a FastTransfer download
 * of an already-open Folder or Message (or Calendar item, itself a Message object on the wire) object,
 * producing a new `"fastTransfer"` Server object handle `RopFastTransferSourceGetBuffer` pages the built stream
 * out of. See `FastTransferStream.ts`'s own doc comment for this pragmatic subset's "full, non-differential
 * dump" ICS scope and its default property columns.
 *
 * `Level` (whether to recurse into subfolders) is decoded to advance the reader correctly but not honored -
 * this pragmatic subset never recurses into subfolders regardless. `CopyFlags`/`SendOptions` (Move mode, best-
 * body preference, Unicode preference, recoverable-mode, ...) are likewise decoded but not honored - this
 * pragmatic subset always copies (never moves) and always encodes the same way regardless of client
 * preference, matching `RopGetPropertiesSpecificHandler`'s own treatment of its analogous `WantUnicode` field.
 * `PropertyTags` is the properties to *exclude* from the default column set (`[MS-OXCROPS]`'s own wording) -
 * see `FastTransferStream.ts` for why an explicit include list instead is `RopFastTransferSourceCopyProperties`'s
 * own job, not this ROP's.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopFastTransferSourceCopyToHandler implements RopHandler {
    public readonly ropId = ROP_ID_COPY_TO;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // Level - subfolder recursion not honored, see class doc comment
        reader.readUInt32LE(); // CopyFlags - not honored, see class doc comment
        reader.readUInt8(); // SendOptions - not honored, see class doc comment
        const propertyTagCount: number = reader.readUInt16LE();
        const excludePropertyIds = new Set<number>();
        for (let i = 0; i < propertyTagCount; i++) {
            excludePropertyIds.add(readPropertyTag(reader).propertyId);
        }

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || (handle.type !== "folder" && handle.type !== "message")) {
            writer.writeUInt8(ROP_ID_COPY_TO);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const transferBuffer = await buildFastTransferStream(handle, context, { excludePropertyIds });
        context.session.handles[outputHandleIndex] = {
            type: "fastTransfer",
            entityUid: handle.entityUid,
            transferBufferBase64: transferBuffer.toString("base64"),
            transferPosition: 0,
        };

        writer.writeUInt8(ROP_ID_COPY_TO);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }
}
