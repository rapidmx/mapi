///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { writeTypedString } from "../codec/TypedString.js";
import { resolveCalendarEventInfo } from "./CalendarEventTarget.js";
import { resolveMessageInfo } from "./MessageTarget.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_OPEN_MESSAGE = 0x03;

/** The well-known MAPI HRESULT `MAPI_E_NOT_FOUND`, reused as this handler's only failure `ReturnValue` - an
 * unrecognized `MessageId`, i.e. not something a `RopQueryRows` `PidTagMid` column has handed out this session
 * (see `MessageTarget.assignOrGetMid`'s own doc comment). Same constant `RopOpenFolderHandler` uses for its
 * own analogous `FolderId` lookup miss. */
const ERROR_NOT_FOUND = 0x8004010f;

/**
 * `RopOpenMessage` (`[MS-OXCMSG]`/`[MS-OXCROPS]`): opens a message by MID, producing a new `"message"` Server
 * object handle later ROPs (`RopGetPropertiesSpecific`, `RopOpenStream`) reference by `InputHandleIndex`. The
 * request's own `FolderId` (the message's expected parent folder) and `CodePageId` are decoded to advance past
 * them correctly but not otherwise acted on - a MID is resolved directly via `session.messageIds`
 * (`MessageTarget.assignOrGetMid`), the same "the session already knows what this small integer refers to"
 * design `RopOpenFolderHandler` uses for FIDs, so cross-checking the caller-supplied `FolderId` would be pure
 * redundancy. `OpenModeFlags` is likewise decoded but not honored - this pragmatic subset only ever opens a
 * message for reading (no `RopSetProperties`/`RopSaveChangesMessage` write-back support yet).
 *
 * **Pragmatic success response**: `HasNamedProperties` is always `false` - a conservative hint value only, not
 * a claim that named properties are unsupported (`RopGetPropertyIdsFromNames`/`PropertyResolvers.ts` do resolve
 * them for a subsequently-opened Calendar item; a real client calls `RopGetPropertyIdsFromNames` unconditionally
 * for a known Calendar message class regardless of this flag, so leaving it `false` doesn't block that path).
 * The `SubjectPrefix`/`NormalizedSubject` `TypedString`s are "no prefix" + the item's plain subject/title (this
 * data model has no separate prefix/normalized-subject split the way real Exchange does), and the recipient
 * table (`RecipientCount`/`ColumnCount`/`RecipientColumns`/`RowCount`/`RecipientRows`) is always empty - reading
 * a message's `To`/`Cc` list is a documented gap in this pass, not silently wrong data.
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopOpenMessageHandler implements RopHandler {
    public readonly ropId = ROP_ID_OPEN_MESSAGE;

    public async handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        reader.readUInt8(); // InputHandleIndex - the folder/logon handle this open is performed against; not
        // separately validated, matching RopOpenFolderHandler's own treatment of its input handle.
        const outputHandleIndex: number = reader.readUInt8();
        reader.readUInt16LE(); // CodePageId - this pragmatic subset always encodes strings as UTF-16LE/UTF-8 directly
        reader.readBigUInt64LE(); // FolderId - the message's expected parent folder; not cross-checked, see class doc comment
        reader.readUInt8(); // OpenModeFlags - read-only opens only in this pragmatic subset
        const messageId: bigint = reader.readBigUInt64LE();

        const target: string | undefined = context.session.messageIds[messageId.toString()];
        if (!target) {
            writer.writeUInt8(ROP_ID_OPEN_MESSAGE);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt32LE(ERROR_NOT_FOUND);
            return;
        }

        // A calendar item's "subject" is its title (PidTagSubject, the same tag a Message uses) - see
        // CalendarEventTarget.ts's own doc comment for why no separate MID-registry mechanism is needed here.
        const subject = target.startsWith("calendarEvent:")
            ? (await resolveCalendarEventInfo(target, context.calendarEventRepo)).title
            : (await resolveMessageInfo(target, context.messageRepo)).subject;
        context.session.handles[outputHandleIndex] = { type: "message", entityUid: target };

        writer.writeUInt8(ROP_ID_OPEN_MESSAGE);
        writer.writeUInt8(outputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt8(0); // HasNamedProperties - conservative fixed value, see class doc comment
        writeTypedString(writer, undefined); // SubjectPrefix - no prefix/normalized-subject split in this data model
        writeTypedString(writer, subject); // NormalizedSubject
        writer.writeUInt16LE(0); // RecipientCount - reading recipients is a documented gap, see class doc comment
        writer.writeUInt16LE(0); // ColumnCount
        writer.writeUInt8(0); // RowCount
    }
}
