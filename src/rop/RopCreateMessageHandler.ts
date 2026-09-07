///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_CREATE_MESSAGE = 0x06;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused as this handler's only failure `ReturnValue` - an
 * unrecognized `FolderId`, the same constant `RopOpenFolderHandler` uses for its own analogous lookup miss. */
const ERROR_NOT_FOUND = 0x8004010f;

/**
 * `RopCreateMessage` (`[MS-OXCMSG]`/`[MS-OXCROPS]`): begins composing a new message in a folder, producing a
 * new `"message"` Server object handle later ROPs (`RopSetProperties`, `RopOpenStream`, `RopSaveChangesMessage`)
 * reference by `InputHandleIndex`. Unlike `RopOpenMessage`'s handle (an already-real, already-saved message),
 * this handle's `entityUid` is deliberately left `""` - there is no real backing `Message` row, and won't be
 * one until `RopSubmitMessage` actually sends it (this pragmatic subset has no separate Drafts-folder
 * persistence step; see `RopSaveChangesMessageHandler`'s own doc comment). `draftFolderUid` remembers the
 * target folder for that eventual send, and `draftProperties` starts empty, filled in by later
 * `RopSetProperties`/`RopWriteStream` calls.
 *
 * `CodePageId` is decoded to advance past it correctly but not otherwise acted on (this pragmatic subset always
 * encodes strings as UTF-16LE/UTF-8 directly, the same treatment `RopOpenMessageHandler` gives it).
 * `AssociatedFlag` (FAI messages - hidden configuration/rules objects, not real mail) is decoded but not
 * honored - this pragmatic subset's compose/send path only ever produces ordinary mail messages.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopCreateMessageHandler implements RopHandler {
    public readonly ropId = ROP_ID_CREATE_MESSAGE;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        reader.readUInt8(); // InputHandleIndex - the logon handle this create is performed against; not
        // separately validated, matching RopOpenFolderHandler's own treatment of its input handle.
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt16LE(); // CodePageId - see class doc comment
        const folderId: bigint = reader.readBigUInt64LE();
        reader.readUInt8(); // AssociatedFlag - see class doc comment

        const target: string | undefined = context.session.folderIds[folderId.toString()];
        if (!target) {
            writer.writeUInt8(ROP_ID_CREATE_MESSAGE);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_NOT_FOUND);
            return;
        }

        context.session.handles[outputHandleIndex] = {
            type: "message",
            entityUid: "",
            draftFolderUid: target,
            draftProperties: {},
        };

        writer.writeUInt8(ROP_ID_CREATE_MESSAGE);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(0); // HasMessageId - no MID is assigned until RopSaveChangesMessage
    }
}
