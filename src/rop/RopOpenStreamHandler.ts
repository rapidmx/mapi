///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { PropertyTag, PropertyType, readPropertyTag } from "../codec/PropertyValue.js";
import { PID_TAG_BODY, resolveMessageBodyBytes } from "./MessageBodyStream.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_OPEN_STREAM = 0x2b;

/** `[MS-OXCPRPT]` §2.2.14.1 `OpenModeFlags` values - `ReadOnly` opens the existing value for reading;
 * `ReadWrite`/`Create` both open for writing (`Create` additionally discards any existing value first, "the
 * mode required for a property that has not been set" per the spec - always true for a fresh
 * `RopCreateMessage` draft's body, which is the only write-mode case this pragmatic subset ever produces). */
const OPEN_MODE_READ_ONLY = 0x00;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused for "the referenced handle isn't a message, or the
 * requested property isn't one this pragmatic subset can stream" - the same constant `RopOpenFolderHandler`
 * uses for its own analogous lookup-miss case. */
const ERROR_NOT_FOUND = 0x8004010f;

/**
 * `RopOpenStream` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): opens a property for streaming access, producing a new
 * `"stream"` Server object handle `RopReadStream`/`RopWriteStream` reads/writes by `InputHandleIndex`.
 * **Pragmatic subset**: only `PidTagBody` (`PtypString`, the plain-text message body) on a `"message"` handle
 * is supported - a real client also streams `PidTagHtml`/`PidTagRtfCompressed` (rich body formats) and
 * folder/attachment binary properties, none of which this pass implements; requesting any other `PropertyTag`,
 * or opening a stream against a non-message handle, fails with `MAPI_E_NOT_FOUND` rather than succeeding with
 * wrong data.
 *
 * `OpenModeFlags` (`[MS-OXCPRPT]` §2.2.14.1) selects which of two independent code paths runs:
 * - `ReadOnly` (`0x00`): the step-7 behavior - resolves the real, already-saved message's body via
 * `MessageBodyStream.resolveMessageBodyBytes()` (`BlobStore` + `mailparser`) and reports its length as
 * `StreamSize`, for `RopReadStream` to read back.
 * - `ReadWrite`/`Create` (`0x01`/`0x02`): a **write-mode** stream, the step-8 addition compose/send needs to
 * accept a message body via `RopWriteStream` - `StreamSize` is always `0` (matching `Create`'s "discards the
 * existing value" semantics; this pragmatic subset never opens `ReadWrite` against a body that already has
 * content) and the new handle's `writeTargetHandleIndex` remembers which message handle to write back into,
 * the link `RopWriteStream`/`RopSaveChangesMessageHandler`/`RopSubmitMessageHandler` use to find it again.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopOpenStreamHandler implements RopHandler {
    public readonly ropId = ROP_ID_OPEN_STREAM;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const outputHandleIndex: number = reader.readUInt8();
        const propertyTag: PropertyTag = readPropertyTag(reader);
        const openModeFlags: number = reader.readUInt8();

        const handle = context.session.handles[inputHandleIndex];
        const isSupportedProperty = propertyTag.propertyId === PID_TAG_BODY && propertyTag.propertyType === PropertyType.PtypString;
        if (!handle || handle.type !== "message" || !isSupportedProperty) {
            writer.writeUInt8(ROP_ID_OPEN_STREAM);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_NOT_FOUND);
            return;
        }

        let streamSize: number;
        if (openModeFlags === OPEN_MODE_READ_ONLY) {
            const bytes = await resolveMessageBodyBytes(handle.entityUid, context.messageRepo, context.blobStore);
            streamSize = bytes.length;
            context.session.handles[outputHandleIndex] = {
                type: "stream",
                entityUid: handle.entityUid,
                propertyId: propertyTag.propertyId,
                propertyType: propertyTag.propertyType,
                streamPosition: 0,
            };
        } else {
            streamSize = 0;
            context.session.handles[outputHandleIndex] = {
                type: "stream",
                entityUid: handle.entityUid,
                propertyId: propertyTag.propertyId,
                propertyType: propertyTag.propertyType,
                writeTargetHandleIndex: inputHandleIndex,
                writeBufferBase64: "",
            };
        }

        writer.writeUInt8(ROP_ID_OPEN_STREAM);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt32LE(streamSize); // StreamSize
    }
}
