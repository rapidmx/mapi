///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { resolveCalendarEventInfo, resolveFolderCalendarEvents } from "../../src/rop/CalendarEventTarget.js";
import { AttendeeResponseStatus, AttendeeRole, BusyStatus, RecurrenceFrequency } from "@rapidmx/restapi";

describe("CalendarEventTarget Tests", () => {
    describe("resolveCalendarEventInfo", () => {
        it("Resolves a real calendar event's fields.", async () => {
            const startDate = new Date("2026-09-07T14:00:00.000Z");
            const endDate = new Date("2026-09-07T15:00:00.000Z");
            const recurrenceRule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const calendarEventRepo = {
                findOne: vi.fn().mockResolvedValue({
                    uid: "evt1",
                    title: "Standup",
                    location: "Room 1",
                    startDate,
                    endDate,
                    timezone: "America/Los_Angeles",
                    busyStatus: BusyStatus.BUSY,
                    recurrenceRule,
                    reminderMinutesBeforeStart: 15,
                    organizer: { address: "organizer@example.com" },
                    attendees: [{ address: "attendee@example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.ACCEPTED, isOrganizer: false }],
                }),
            };

            const info = await resolveCalendarEventInfo("calendarEvent:evt1", calendarEventRepo as any);

            expect(info).toEqual({
                title: "Standup",
                location: "Room 1",
                startDate,
                endDate,
                timezone: "America/Los_Angeles",
                busyStatus: BusyStatus.BUSY,
                recurrenceRule,
                reminderMinutesBeforeStart: 15,
                organizerAddress: "organizer@example.com",
                attendees: [{ address: "attendee@example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.ACCEPTED, isOrganizer: false }],
            });
            expect(calendarEventRepo.findOne).toHaveBeenCalledWith("evt1", { ignoreACL: true });
        });

        it("Degrades to empty-looking values for a calendarEvent target whose real CalendarEvent has since vanished.", async () => {
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(undefined) };

            const info = await resolveCalendarEventInfo("calendarEvent:gone", calendarEventRepo as any);

            expect(info.title).toBe("");
            expect(info.location).toBeUndefined();
            expect(info.startDate).toEqual(new Date(0));
            expect(info.endDate).toEqual(new Date(0));
            expect(info.timezone).toBe("UTC");
            expect(info.busyStatus).toBe(BusyStatus.BUSY);
            expect(info.recurrenceRule).toBeUndefined();
            expect(info.reminderMinutesBeforeStart).toBeUndefined();
            expect(info.organizerAddress).toBe("");
            expect(info.attendees).toEqual([]);
        });
    });

    describe("resolveFolderCalendarEvents", () => {
        it("Returns each event as a calendarEvent:<uid> target string.", async () => {
            const calendarEventRepo = { find: vi.fn().mockResolvedValue([{ uid: "evt1" }, { uid: "evt2" }]) };

            const targets = await resolveFolderCalendarEvents("folder-1", calendarEventRepo as any);

            expect(targets).toEqual(["calendarEvent:evt1", "calendarEvent:evt2"]);
            expect(calendarEventRepo.find).toHaveBeenCalledWith({ folderUid: "folder-1" }, { ignoreACL: true });
        });
    });
});
