///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { readPropertyTag, type PropertyTag } from "../codec/PropertyValue.js";
import { buildFastTransferStream } from "./FastTransferStream.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_COPY_PROPERTIES = 0x69;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a folder or
 * message (or doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its own analogous
 * check. */
const ERROR_INVALID_OBJECT = 0x80070005;

/**
 * `RopFastTransferSourceCopyProperties` (`[MS-OXCFXICS]`/`[MS-OXCROPS]`, RopId `0x69`): the explicit-column
 * sibling of `RopFastTransferSourceCopyTo` - `PropertyTags` here is the properties to **copy** (an include
 * list, per `[MS-OXCROPS]`'s own wording, the opposite of `CopyTo`'s exclude list), applied uniformly at both
 * folder- and message-level wherever `FastTransferStream.ts` builds a `propList` - see that file's own doc
 * comment for this pragmatic subset's overall ICS scope. An empty `PropertyTags` list is honored literally
 * (copies zero properties, i.e. empty `propList`s) rather than falling back to `CopyTo`'s own default column
 * set - that fallback belongs to `CopyTo` alone, since an explicit "copy nothing" request here is spec-valid
 * and distinguishable from "use the default set" (`CopyTo` simply has no `PropertyTags`-supplied alternative to
 * fall back from).
 *
 * `Level`/`CopyFlags`/`SendOptions` are decoded to advance the reader correctly but not honored - see
 * `RopFastTransferSourceCopyToHandler`'s own doc comment for the identical reasoning. Note `CopyFlags` here is
 * **1 byte**, not 4 like `CopyTo`'s own field of the same name - confirmed against each ROP's own request-buffer
 * page separately, not assumed identical.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopFastTransferSourceCopyPropertiesHandler implements RopHandler {
    public readonly ropId = ROP_ID_COPY_PROPERTIES;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt8(); // Level - subfolder recursion not honored, see class doc comment
        reader.readUInt8(); // CopyFlags (1 byte here, unlike CopyTo's 4) - not honored, see class doc comment
        reader.readUInt8(); // SendOptions - not honored, see class doc comment
        const propertyTagCount: number = reader.readUInt16LE();
        const columns: PropertyTag[] = [];
        for (let i = 0; i < propertyTagCount; i++) {
            columns.push(readPropertyTag(reader));
        }

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || (handle.type !== "folder" && handle.type !== "message")) {
            writer.writeUInt8(ROP_ID_COPY_PROPERTIES);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        const transferBuffer = await buildFastTransferStream(handle, context, { columns });
        context.session.handles[outputHandleIndex] = {
            type: "fastTransfer",
            entityUid: handle.entityUid,
            transferBufferBase64: transferBuffer.toString("base64"),
            transferPosition: 0,
        };

        writer.writeUInt8(ROP_ID_COPY_PROPERTIES);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
    }
}
