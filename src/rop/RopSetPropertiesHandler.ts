///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { PropertyType, readTaggedPropertyValue, type TaggedPropertyValue } from "../codec/PropertyValue.js";
import {
    LID_APPOINTMENT_END_WHOLE,
    LID_APPOINTMENT_RECUR,
    LID_APPOINTMENT_START_WHOLE,
    LID_BUSY_STATUS,
    LID_GLOBAL_OBJECT_ID,
    LID_LOCATION,
    LID_RECURRING,
    LID_REMINDER_DELTA,
    LID_REMINDER_SET,
    LID_TIME_ZONE_STRUCT,
    PSETID_APPOINTMENT,
    PSETID_COMMON,
    PSETID_MEETING,
} from "./CalendarNamedProperties.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import type { RopContext, RopHandler } from "./RopHandler.js";

const ROP_ID_SET_PROPERTIES = 0x0a;

/** The well-known MAPI HRESULT `MAPI_E_INVALID_OBJECT`, reused for "the referenced handle isn't a message (or
 * doesn't exist)" - the same constant `RopGetPropertiesSpecificHandler` uses for its own analogous check.
 * `RopSetProperties` is also spec-valid on Folder/Attachment/Logon objects, none of which this pragmatic
 * subset's compose/send path ever needs to set properties on. */
const ERROR_INVALID_OBJECT = 0x80070005;

// The small, well-known set of plain (non-named) properties this pragmatic subset's compose/send path tracks -
// exactly what RopSubmitMessageHandler/RopSaveChangesMessageHandler need to build a MIME message or persist a
// CalendarEvent (Subject, an inline small body, the PidTagDisplayTo/Cc/Bcc "cached recipient display string"
// properties real Outlook also always sets alongside RopModifyRecipients - see RopSubmitMessageHandler's own
// doc comment for why this pragmatic subset reads addresses from these instead of implementing
// RopModifyRecipients - and PidTagMessageClass, which RopSaveChangesMessageHandler/RopSubmitMessageHandler
// branch on to route Calendar items down a different path than ordinary mail). Every other plain property a
// client sets is accepted (this ROP always reports success, no PropertyProblems) but simply not tracked -
// harmless for this pragmatic subset's narrow compose/send scope. A *named* property (ID `>= 0x8000`, resolved
// via `resolveNamedProperty`) is tracked separately below, against the Calendar LID table `PropertyResolvers.ts`
// also reads from - see `CalendarNamedProperties.ts`.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_CLASS = 0x001a;
const PID_TAG_DISPLAY_BCC = 0x0e02;
const PID_TAG_DISPLAY_CC = 0x0e03;
const PID_TAG_DISPLAY_TO = 0x0e04;
const PID_TAG_BODY = 0x1000;
const TRACKED_PROPERTY_IDS: ReadonlySet<number> = new Set([
    PID_TAG_SUBJECT,
    PID_TAG_MESSAGE_CLASS,
    PID_TAG_DISPLAY_BCC,
    PID_TAG_DISPLAY_CC,
    PID_TAG_DISPLAY_TO,
    PID_TAG_BODY,
]);

/** The Calendar named-property LIDs this handler tracks when set under the correct property-set GUID (per
 * `CalendarNamedProperties.ts`) - the write-side mirror of `PropertyResolvers.calendarEventValueFor`'s read-side
 * switch. `PidLidRecurring`/`PidLidReminderSet` are deliberately not tracked here: both are purely-derived
 * boolean flags on the read side (`recurrenceRule !== undefined` / `reminderMinutesBeforeStart != null`), so a
 * client setting them carries no additional information `RopSaveChangesMessageHandler` needs to persist. */
const TRACKED_APPOINTMENT_LIDS: ReadonlySet<number> = new Set([
    LID_LOCATION,
    LID_APPOINTMENT_START_WHOLE,
    LID_APPOINTMENT_END_WHOLE,
    LID_BUSY_STATUS,
    LID_APPOINTMENT_RECUR,
    LID_TIME_ZONE_STRUCT,
]);
const TRACKED_COMMON_LIDS: ReadonlySet<number> = new Set([LID_REMINDER_DELTA]);

/** `PidLidGlobalObjectId` - the meeting-response side needs this tracked so `RopSubmitMessageHandler` can
 * correlate an incoming `"IPM.Schedule.Meeting.Resp.*"` response back to the `CalendarEvent` it answers, via
 * `MeetingMessageClassHandler.ts`. */
const TRACKED_MEETING_LIDS: ReadonlySet<number> = new Set([LID_GLOBAL_OBJECT_ID]);

/**
 * `RopSetProperties` (`[MS-OXCPRPT]`/`[MS-OXCROPS]`): sets an explicit, client-chosen list of properties
 * (`TaggedPropertyValue`s, each carrying its own `PropertyTag`) on an already-open object. This pragmatic
 * subset supports only `"message"` handles (a `RopCreateMessage` draft, or an already-open `RopOpenMessage`
 * handle) and only tracks the small property set `TRACKED_PROPERTY_IDS`/`TRACKED_APPOINTMENT_LIDS`/
 * `TRACKED_COMMON_LIDS` name - see those constants' own comments for why. A named property (ID `>= 0x8000`) is
 * resolved back to its `(PropertySet GUID, LID)` identity via `resolveNamedProperty` before it can be tracked,
 * since MAPI has no fixed numeric ID for any Calendar property the way `PidTagSubject` has one.
 *
 * Every value is coerced to a plain string via `stringifyValue` before being stored in `handle.draftProperties`
 * (itself a `Record<string, string>` for `MapiSessionContext`'s own JSON-safety reasons, see
 * `MapiSessionManager.ts`) - a lossless, type-aware encoding (ISO-8601 for `PtypTime`, base64 for `PtypBinary`,
 * plain `String()` otherwise) so `RopSaveChangesMessageHandler` can decode a Calendar item's properties back
 * into a real `CalendarEvent`'s typed fields, not just display strings.
 *
 * Always reports success with zero `PropertyProblems`, even for a property this pragmatic subset doesn't
 * track - the same "never fail over an unsupported property" pragmatic stance `RopQueryRows`/
 * `RopGetPropertiesSpecific` already take (falling back to a default value there; simply not tracking the
 * property here, since there is no row to fill in for a *set* operation).
 *
 * @author Jean-Philippe Steinmetz
 */
export class RopSetPropertiesHandler implements RopHandler {
    public readonly ropId = ROP_ID_SET_PROPERTIES;

    public handle(reader: BufferReader, writer: BufferWriter, context: RopContext): void {
        reader.readUInt8(); // LogonId - this pragmatic subset doesn't track multiple concurrent logons per session
        const inputHandleIndex: number = reader.readUInt8();
        const propertyValueSize: number = reader.readUInt16LE();
        // PropertyValueSize covers the PropertyValueCount field itself plus PropertyValues, counted from right
        // here (immediately after the PropertyValueSize field, before PropertyValueCount is read).
        const propertyValuesEnd: number = reader.position + propertyValueSize;
        const propertyValueCount: number = reader.readUInt16LE();
        const values: TaggedPropertyValue[] = [];
        for (let i = 0; i < propertyValueCount; i++) {
            values.push(readTaggedPropertyValue(reader));
        }
        // PropertyValueSize is a byte-count check the request buffer itself provides for framing purposes (so a
        // generic ROP-skipping implementation could skip this ROP without decoding it) - this handler already
        // decoded PropertyValues field-by-field above, so the only remaining use is confirming the reader ended
        // up exactly where PropertyValueSize said it would, catching a malformed request loudly rather than
        // silently misaligning every ROP that follows in the same RopsList.
        if (reader.position !== propertyValuesEnd) {
            throw new Error("RopSetProperties: PropertyValueSize did not match the decoded PropertyValues length.");
        }

        const handle = context.session.handles[inputHandleIndex];
        if (!handle || handle.type !== "message") {
            writer.writeUInt8(ROP_ID_SET_PROPERTIES);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt32LE(ERROR_INVALID_OBJECT);
            return;
        }

        handle.draftProperties ??= {};
        for (const tagged of values) {
            if (TRACKED_PROPERTY_IDS.has(tagged.propertyId)) {
                handle.draftProperties[String(tagged.propertyId)] = stringifyValue(tagged);
            } else if (tagged.propertyId >= 0x8000 && this.isTrackedNamedProperty(context, tagged.propertyId)) {
                handle.draftProperties[String(tagged.propertyId)] = stringifyValue(tagged);
            }
        }

        writer.writeUInt8(ROP_ID_SET_PROPERTIES);
        writer.writeUInt8(inputHandleIndex);
        writer.writeUInt32LE(0); // ReturnValue - success
        writer.writeUInt16LE(0); // PropertyProblemCount - never reported, see class doc comment
    }

    private isTrackedNamedProperty(context: RopContext, propertyId: number): boolean {
        const named = resolveNamedProperty(context.session, propertyId);
        if (!named || named.kind !== "lid" || named.lid === undefined) {
            return false;
        }
        const guid = named.guid.toLowerCase();
        if (guid === PSETID_APPOINTMENT) {
            return TRACKED_APPOINTMENT_LIDS.has(named.lid);
        }
        if (guid === PSETID_COMMON) {
            return TRACKED_COMMON_LIDS.has(named.lid);
        }
        if (guid === PSETID_MEETING) {
            return TRACKED_MEETING_LIDS.has(named.lid);
        }
        return false;
    }
}

function stringifyValue(tagged: TaggedPropertyValue): string {
    switch (tagged.propertyType) {
        case PropertyType.PtypString:
        case PropertyType.PtypString8:
            return tagged.value as string;
        case PropertyType.PtypTime:
            return (tagged.value as Date).toISOString();
        case PropertyType.PtypBinary:
            return (tagged.value as Buffer).toString("base64");
        default:
            return String(tagged.value);
    }
}
