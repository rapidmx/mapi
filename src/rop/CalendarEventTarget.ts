///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { Attendee, BusyStatus, CalendarEvent, RecurrenceRule } from "@rapidmx/restapi";

/**
 * The `RopGetContentsTable`/`RopOpenMessage` analog of `MessageTarget.ts`, for a `Folder` of type `CALENDAR`:
 * resolves a `"calendarEvent:<uid>"` row/handle target into the display data a calendar-item row or an opened
 * appointment's properties need. A `CalendarEvent` is addressed by MID exactly the same way a `Message` is -
 * `MessageTarget.assignOrGetMid`/`session.messageIds` are reused as-is (see that file's own doc comment; the
 * mechanism is already generic over the target-string prefix, not message-specific despite its file location),
 * so this file adds no MID-registry code of its own.
 */
export interface CalendarEventTargetInfo {
    title: string;
    location?: string;
    startDate: Date;
    endDate: Date;
    timezone: string;
    busyStatus: BusyStatus;
    recurrenceRule?: RecurrenceRule;
    reminderMinutesBeforeStart?: number;
    organizerAddress: string;
    attendees: Attendee[];
}

/** Degrades to empty-looking values for a `"calendarEvent:<uid>"` target whose real `CalendarEvent` has since
 * vanished (soft-deleted or otherwise) - the same "don't fail the whole ROP over one stale row" principle
 * `MessageTarget.resolveMessageInfo`/`FolderTarget.resolveFolderInfo` already apply. */
export async function resolveCalendarEventInfo(target: string, calendarEventRepo: RepoUtils<any>): Promise<CalendarEventTargetInfo> {
    const uid = target.slice("calendarEvent:".length);
    const event: CalendarEvent | undefined = await calendarEventRepo.findOne(uid, { ignoreACL: true });
    return {
        title: event?.title ?? "",
        location: event?.location,
        startDate: event?.startDate ?? new Date(0),
        endDate: event?.endDate ?? new Date(0),
        timezone: event?.timezone ?? "UTC",
        busyStatus: event?.busyStatus ?? BusyStatus.BUSY,
        recurrenceRule: event?.recurrenceRule,
        reminderMinutesBeforeStart: event?.reminderMinutesBeforeStart,
        organizerAddress: event?.organizer?.address ?? "",
        attendees: event?.attendees ?? [],
    };
}

/** Resolves the calendar events directly in `folderUid`, as `"calendarEvent:<uid>"` target strings, for
 * `RopGetContentsTable` against a `CALENDAR`-type folder. */
export async function resolveFolderCalendarEvents(folderUid: string, calendarEventRepo: RepoUtils<any>): Promise<string[]> {
    const events: CalendarEvent[] = await calendarEventRepo.find({ folderUid }, { ignoreACL: true });
    return events.map((e) => `calendarEvent:${e.uid}`);
}
