///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { encodeAppointmentRecurrence } from "../codec/AppointmentRecurrence.js";
import { encodeTimeZoneStruct } from "../codec/MapiTimeZone.js";
import { PropertyType, PropertyValueData } from "../codec/PropertyValue.js";
import type { MapiSessionContext } from "../MapiSessionManager.js";
import {
    BUSY_STATUS_CODES,
    LID_APPOINTMENT_END_WHOLE,
    LID_APPOINTMENT_RECUR,
    LID_APPOINTMENT_START_WHOLE,
    LID_BUSY_STATUS,
    LID_LOCATION,
    LID_RECURRING,
    LID_REMINDER_DELTA,
    LID_REMINDER_SET,
    LID_RESPONSE_STATUS,
    LID_TIME_ZONE_STRUCT,
    PSETID_APPOINTMENT,
    PSETID_COMMON,
    RESPONSE_STATUS_CODES,
    RESPONSE_STATUS_NONE,
    RESPONSE_STATUS_ORGANIZED,
} from "./CalendarNamedProperties.js";
import { CalendarEventTargetInfo, resolveCalendarEventInfo } from "./CalendarEventTarget.js";
import { assignOrGetFid, FolderTargetInfo, resolveFolderInfo } from "./FolderTarget.js";
import { assignOrGetMid, MessageTargetInfo, resolveMessageInfo } from "./MessageTarget.js";
import { resolveNamedProperty } from "./NamedPropertyRegistry.js";
import type { RopContext } from "./RopHandler.js";

// Well-known folder property IDs this pragmatic subset supports - the small set a real client needs to render
// a folder-hierarchy view. Add more as a real need arises, not speculatively.
const PID_TAG_DISPLAY_NAME = 0x3001;
const PID_TAG_FOLDER_ID = 0x6748;
const PID_TAG_CONTENT_COUNT = 0x3602;
const PID_TAG_CONTENT_UNREAD_COUNT = 0x3603;
const PID_TAG_SUBFOLDERS = 0x360a;

// Well-known message property IDs this pragmatic subset supports - the small set a real client needs to render
// a message list or a single message's metadata.
const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MESSAGE_FLAGS = 0x0e07;
const PID_TAG_HAS_ATTACHMENTS = 0x0e1b;
const PID_TAG_MESSAGE_DELIVERY_TIME = 0x0e06;
const PID_TAG_MID = 0x674a;
/** `MSGFLAG_READ`, the one `PidTagMessageFlags` bit this pragmatic subset ever sets. */
const MSGFLAG_READ = 0x01;

// Calendar named-property identity (`(PropertySet GUID, LID)` pairs) this pragmatic subset supports - see
// CalendarNamedProperties.ts's own doc comment. A calendar item's property IDs `>= 0x8000` are resolved back
// to one of these via `NamedPropertyRegistry.resolveNamedProperty` before this switch can run, since MAPI has
// no fixed numeric ID for any of them the way `PidTagSubject` has one.

/** A type-appropriate zero/empty value for a requested property this handler has no real data for - keeps
 * `StandardPropertyRow` encoding valid (a real value of the right type, per `writePropertyValue`'s
 * expectations) without needing to model every property a client could ever ask for. Shared by
 * `RopQueryRowsHandler` (table rows) and `RopGetPropertiesSpecificHandler` (single-object property fetch). */
export function defaultValueForType(propertyType: PropertyType): PropertyValueData {
    switch (propertyType) {
        case PropertyType.PtypBoolean:
            return false;
        case PropertyType.PtypInteger16:
        case PropertyType.PtypInteger32:
        case PropertyType.PtypFloating32:
        case PropertyType.PtypFloating64:
            return 0;
        case PropertyType.PtypInteger64:
            return 0n;
        case PropertyType.PtypTime:
            return new Date(0);
        case PropertyType.PtypGuid:
            return "00000000-0000-0000-0000-000000000000";
        case PropertyType.PtypBinary:
            return Buffer.alloc(0);
        case PropertyType.PtypMultipleInteger32:
        case PropertyType.PtypMultipleString:
        case PropertyType.PtypMultipleString8:
        case PropertyType.PtypMultipleBinary:
            return [];
        case PropertyType.PtypString:
        case PropertyType.PtypString8:
        default:
            return "";
    }
}

/** Resolves one requested property's value for a `"folder:<uid>"`/`"virtual:<name>"` target. */
export function folderValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: FolderTargetInfo,
): PropertyValueData {
    switch (propertyId) {
        case PID_TAG_DISPLAY_NAME:
            return info.displayName;
        case PID_TAG_FOLDER_ID:
            return BigInt(assignOrGetFid(session, target));
        case PID_TAG_CONTENT_COUNT:
            return info.totalCount;
        case PID_TAG_CONTENT_UNREAD_COUNT:
            return info.unreadCount;
        case PID_TAG_SUBFOLDERS:
            return info.hasChildren;
        default:
            return defaultValueForType(propertyType);
    }
}

/** Resolves one requested property's value for a `"message:<uid>"` target. */
export function messageValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: MessageTargetInfo,
): PropertyValueData {
    switch (propertyId) {
        case PID_TAG_SUBJECT:
            return info.subject;
        case PID_TAG_MESSAGE_FLAGS:
            return info.read ? MSGFLAG_READ : 0;
        case PID_TAG_HAS_ATTACHMENTS:
            return info.hasAttachments;
        case PID_TAG_MESSAGE_DELIVERY_TIME:
            return info.receivedDate;
        case PID_TAG_MID:
            return BigInt(assignOrGetMid(session, target));
        default:
            return defaultValueForType(propertyType);
    }
}

/** Resolves one requested property's value for a `"calendarEvent:<uid>"` target. Almost every Appointment
 * property is a *named* property (`PidLid*`, no fixed numeric ID) rather than a plain `PidTag*` - a property ID
 * `>= 0x8000` is resolved back to its `(PropertySet GUID, LID)` identity via `resolveNamedProperty` before this
 * can dispatch on it (see `NamedPropertyRegistry.ts`'s own doc comment and the architecture plan's "Calendar
 * support" section for the full named-property table this switch implements). */
export function calendarEventValueFor(
    session: MapiSessionContext,
    propertyId: number,
    propertyType: PropertyType,
    target: string,
    info: CalendarEventTargetInfo,
    callerAddress: string,
): PropertyValueData {
    if (propertyId === PID_TAG_SUBJECT) {
        return info.title;
    }
    if (propertyId === PID_TAG_MID) {
        return BigInt(assignOrGetMid(session, target));
    }
    if (propertyId < 0x8000) {
        return defaultValueForType(propertyType);
    }

    const namedProperty = resolveNamedProperty(session, propertyId);
    if (!namedProperty || namedProperty.kind !== "lid") {
        return defaultValueForType(propertyType);
    }
    const guid = namedProperty.guid.toLowerCase();
    const isAppointmentProperty = guid === PSETID_APPOINTMENT;
    const isCommonProperty = guid === PSETID_COMMON;

    switch (namedProperty.lid) {
        case LID_LOCATION:
            return isAppointmentProperty ? (info.location ?? "") : defaultValueForType(propertyType);
        case LID_APPOINTMENT_START_WHOLE:
            return isAppointmentProperty ? info.startDate : defaultValueForType(propertyType);
        case LID_APPOINTMENT_END_WHOLE:
            return isAppointmentProperty ? info.endDate : defaultValueForType(propertyType);
        case LID_BUSY_STATUS:
            return isAppointmentProperty ? BUSY_STATUS_CODES[info.busyStatus] : defaultValueForType(propertyType);
        case LID_RECURRING:
            return isAppointmentProperty ? info.recurrenceRule !== undefined : defaultValueForType(propertyType);
        case LID_APPOINTMENT_RECUR:
            return isAppointmentProperty && info.recurrenceRule
                ? encodeAppointmentRecurrence(info.recurrenceRule, info.startDate, info.endDate)
                : defaultValueForType(propertyType);
        case LID_TIME_ZONE_STRUCT:
            return isAppointmentProperty ? encodeTimeZoneStruct(info.timezone, info.startDate) : defaultValueForType(propertyType);
        case LID_RESPONSE_STATUS:
            return isAppointmentProperty ? responseStatusFor(info, callerAddress) : defaultValueForType(propertyType);
        case LID_REMINDER_SET:
            return isCommonProperty ? info.reminderMinutesBeforeStart != null : defaultValueForType(propertyType);
        case LID_REMINDER_DELTA:
            return isCommonProperty ? (info.reminderMinutesBeforeStart ?? 0) : defaultValueForType(propertyType);
        default:
            return defaultValueForType(propertyType);
    }
}

/** The caller's own `PidLidResponseStatus` value: `respOrganized` if the caller mailbox *is* this event's
 * organizer, else the matching `Attendee.responseStatus` (by address, case-insensitive), else `respNone` if the
 * caller isn't party to this event at all (e.g. a shared/delegate calendar view). */
function responseStatusFor(info: CalendarEventTargetInfo, callerAddress: string): number {
    if (callerAddress && info.organizerAddress.toLowerCase() === callerAddress.toLowerCase()) {
        return RESPONSE_STATUS_ORGANIZED;
    }
    const attendee = info.attendees.find((a) => a.address.toLowerCase() === callerAddress.toLowerCase());
    return attendee ? RESPONSE_STATUS_CODES[attendee.responseStatus] : RESPONSE_STATUS_NONE;
}

/** Resolves every column in `columns` for a single `target` (a `"folder:"`/`"virtual:"`/`"message:"`/
 * `"calendarEvent:"` target string), in order - the shared implementation behind both `RopQueryRowsHandler`
 * (one call per table row) and `RopGetPropertiesSpecificHandler` (one call for the single object a handle
 * refers to). */
export async function resolvePropertyValues(
    target: string,
    columns: { propertyId: number; propertyType: PropertyType }[],
    context: Pick<RopContext, "mailboxUid" | "session" | "folderRepo" | "messageRepo" | "calendarEventRepo" | "mailboxRepo">,
): Promise<PropertyValueData[]> {
    const isMessage = target.startsWith("message:");
    const isCalendarEvent = target.startsWith("calendarEvent:");
    const folderInfo =
        isMessage || isCalendarEvent ? undefined : await resolveFolderInfo(context.mailboxUid, target, context.folderRepo);
    const messageInfo = isMessage ? await resolveMessageInfo(target, context.messageRepo) : undefined;
    const calendarEventInfo = isCalendarEvent ? await resolveCalendarEventInfo(target, context.calendarEventRepo) : undefined;

    let callerAddress = "";
    if (isCalendarEvent) {
        const mailbox = await context.mailboxRepo.findOne(context.mailboxUid, { ignoreACL: true });
        callerAddress = mailbox?.primarySmtpAddress ?? "";
    }

    return columns.map((column) => {
        if (calendarEventInfo) {
            return calendarEventValueFor(context.session, column.propertyId, column.propertyType, target, calendarEventInfo, callerAddress);
        }
        if (messageInfo) {
            return messageValueFor(context.session, column.propertyId, column.propertyType, target, messageInfo);
        }
        return folderValueFor(context.session, column.propertyId, column.propertyType, target, folderInfo!);
    });
}
