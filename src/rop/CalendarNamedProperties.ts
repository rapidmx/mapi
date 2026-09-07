///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { AttendeeResponseStatus, BusyStatus } from "@rapidmx/restapi";

/**
 * The Calendar named-property identity table (`(PropertySet GUID, LID)` pairs) this pragmatic subset supports,
 * per `[MS-OXPROPS]` - each pinned against its own MS-OXPROPS page (not invented), shared by `PropertyResolvers.ts`
 * (the read side: `calendarEventValueFor`) and `RopSaveChangesMessageHandler.ts` (the write side: decoding a
 * draft's accumulated `RopSetProperties` values back into a real `CalendarEvent`) so both directions of this
 * mapping stay in exactly one place. See the architecture plan's "Calendar support" section for the full table.
 */
export const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
export const PSETID_COMMON = "00062008-0000-0000-c000-000000000046";
export const PSETID_MEETING = "6ed8da90-450b-101b-98da-00aa003f1305";

export const LID_LOCATION = 0x8208;
export const LID_APPOINTMENT_START_WHOLE = 0x820d;
export const LID_APPOINTMENT_END_WHOLE = 0x820e;
export const LID_BUSY_STATUS = 0x8205;
export const LID_RECURRING = 0x8223;
export const LID_REMINDER_SET = 0x8503;
export const LID_REMINDER_DELTA = 0x8501;
export const LID_RESPONSE_STATUS = 0x8218;
export const LID_APPOINTMENT_RECUR = 0x8216;
export const LID_TIME_ZONE_STRUCT = 0x8233;

/** `PidLidGlobalObjectId` (`PSETID_Meeting`) - see `GlobalObjectId.ts` for the BLOB it carries and
 * `MeetingMessageClassHandler.ts` for how a meeting response's own copy of it is correlated back to the
 * original `CalendarEvent`. */
export const LID_GLOBAL_OBJECT_ID = 0x00000003;

/** `PidLidBusyStatus`'s wire values (`[MS-OXOCAL]` §2.2.1.2, confirmed this session) - a different numbering
 * than `CalendarSyncAdapter`'s own `BUSY_STATUS_CODES` table for the unrelated MS-ASCAL `BusyStatus` field.
 * `olWorkingElsewhere` (0x00000004) has no equivalent in this library's own `BusyStatus` enum - never produced,
 * and decoded (via `BUSY_STATUS_FROM_CODE`) with a fallback to `BUSY`, a documented, harmless approximation. */
export const BUSY_STATUS_CODES: Record<BusyStatus, number> = {
    [BusyStatus.FREE]: 0,
    [BusyStatus.TENTATIVE]: 1,
    [BusyStatus.BUSY]: 2,
    [BusyStatus.OUT_OF_OFFICE]: 3,
};

export const BUSY_STATUS_FROM_CODE: Record<number, BusyStatus> = {
    0: BusyStatus.FREE,
    1: BusyStatus.TENTATIVE,
    2: BusyStatus.BUSY,
    3: BusyStatus.OUT_OF_OFFICE,
};

/** `PidLidResponseStatus`'s wire values (`[MS-OXOCAL]` §2.2.1.11, confirmed this session) for the caller's own
 * attendee record. `respOrganized` (0x00000001) is returned when the caller *is* the organizer - not a value
 * `AttendeeResponseStatus` itself has a case for, since that enum only models an attendee's own response. */
export const RESPONSE_STATUS_CODES: Record<AttendeeResponseStatus, number> = {
    [AttendeeResponseStatus.NEEDS_ACTION]: 5, // respNotResponded
    [AttendeeResponseStatus.TENTATIVE]: 2, // respTentative
    [AttendeeResponseStatus.ACCEPTED]: 3, // respAccepted
    [AttendeeResponseStatus.DECLINED]: 4, // respDeclined
};
export const RESPONSE_STATUS_ORGANIZED = 1; // respOrganized
export const RESPONSE_STATUS_NONE = 0; // respNone
