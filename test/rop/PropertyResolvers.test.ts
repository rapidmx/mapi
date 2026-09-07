///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { encodeAppointmentRecurrence } from "../../src/codec/AppointmentRecurrence.js";
import { encodeTimeZoneStruct } from "../../src/codec/MapiTimeZone.js";
import { PropertyType } from "../../src/codec/PropertyValue.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { calendarEventValueFor, resolvePropertyValues } from "../../src/rop/PropertyResolvers.js";
import type { CalendarEventTargetInfo } from "../../src/rop/CalendarEventTarget.js";
import {
    AttendeeResponseStatus,
    AttendeeRole,
    BusyStatus,
    RecurrenceFrequency,
    type Attendee,
} from "@rapidmx/restapi";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
const PSETID_COMMON = "00062008-0000-0000-c000-000000000046";

const LID_LOCATION = 0x8208;
const LID_APPOINTMENT_START_WHOLE = 0x820d;
const LID_APPOINTMENT_END_WHOLE = 0x820e;
const LID_BUSY_STATUS = 0x8205;
const LID_RECURRING = 0x8223;
const LID_REMINDER_SET = 0x8503;
const LID_REMINDER_DELTA = 0x8501;
const LID_RESPONSE_STATUS = 0x8218;
const LID_APPOINTMENT_RECUR = 0x8216;
const LID_TIME_ZONE_STRUCT = 0x8233;

const PID_TAG_SUBJECT = 0x0037;
const PID_TAG_MID = 0x674a;

function makeSession(): MapiSessionContext {
    return new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
}

function baseInfo(overrides: Partial<CalendarEventTargetInfo> = {}): CalendarEventTargetInfo {
    return {
        title: "Standup",
        location: "Room 1",
        startDate: new Date("2026-09-07T14:00:00.000Z"),
        endDate: new Date("2026-09-07T15:00:00.000Z"),
        timezone: "UTC",
        busyStatus: BusyStatus.BUSY,
        recurrenceRule: undefined,
        reminderMinutesBeforeStart: undefined,
        organizerAddress: "organizer@example.com",
        attendees: [],
        ...overrides,
    };
}

function namedId(session: MapiSessionContext, guid: string, lid: number): number {
    return assignOrGetNamedPropertyId(session, { guid, kind: "lid", lid });
}

describe("PropertyResolvers Tests", () => {
    describe("calendarEventValueFor", () => {
        it("Resolves PidTagSubject to the event's title.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, PID_TAG_SUBJECT, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("Standup");
        });

        it("Resolves PidTagMid via assignOrGetMid.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, PID_TAG_MID, PropertyType.PtypInteger64, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(1n);
            expect(session.messageIds["1"]).toBe("calendarEvent:e1");
        });

        it("Falls back to a type-appropriate default for an unrecognized plain (< 0x8000) property tag.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, 0x0e07, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(0);
        });

        it("Falls back to a default when the property ID was never assigned by RopGetPropertyIdsFromNames.", () => {
            const session = makeSession();
            const value = calendarEventValueFor(session, 0x8000, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a default for a Kind=name named property (this pragmatic subset only maps Kind=LID Appointment/Common properties).", () => {
            const session = makeSession();
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "name", name: "SomeCustomProp" });
            const value = calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a default for a recognized LID under the wrong property-set GUID.", () => {
            const session = makeSession();
            const id = namedId(session, "11111111-0000-0000-c000-000000000046", LID_LOCATION);
            const value = calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe("");
        });

        it("Falls back to a type-appropriate default for an unrecognized LID under a known property set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, 0x9999);
            const value = calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "");
            expect(value).toBe(0);
        });

        it("Resolves PidLidLocation, defaulting to empty string when unset.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_LOCATION);
            expect(calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo({ location: "Room 9" }), "")).toBe("Room 9");
            expect(calendarEventValueFor(session, id, PropertyType.PtypString, "calendarEvent:e1", baseInfo({ location: undefined }), "")).toBe("");
        });

        it("Resolves PidLidAppointmentStartWhole/EndWhole.", () => {
            const session = makeSession();
            const info = baseInfo();
            const startId = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_START_WHOLE);
            const endId = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_END_WHOLE);
            expect(calendarEventValueFor(session, startId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toBe(info.startDate);
            expect(calendarEventValueFor(session, endId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toBe(info.endDate);
        });

        it("Resolves PidLidBusyStatus to the correct MS-OXOCAL wire value for each BusyStatus.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_BUSY_STATUS);
            const cases: [BusyStatus, number][] = [
                [BusyStatus.FREE, 0],
                [BusyStatus.TENTATIVE, 1],
                [BusyStatus.BUSY, 2],
                [BusyStatus.OUT_OF_OFFICE, 3],
            ];
            for (const [busyStatus, expected] of cases) {
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo({ busyStatus }), "")).toBe(expected);
            }
        });

        it("Resolves PidLidRecurring to whether recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_RECURRING);
            const rule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            expect(calendarEventValueFor(session, id, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo({ recurrenceRule: rule }), "")).toBe(true);
            expect(calendarEventValueFor(session, id, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo({ recurrenceRule: undefined }), "")).toBe(false);
        });

        it("Resolves PidLidAppointmentRecur to the encoded recurrence blob when a recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_RECUR);
            const rule = { freq: RecurrenceFrequency.DAILY, interval: 1, exceptions: [] };
            const info = baseInfo({ recurrenceRule: rule });
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", info, "");
            expect(value).toEqual(encodeAppointmentRecurrence(rule, info.startDate, info.endDate));
        });

        it("Falls back to a default for PidLidAppointmentRecur when no recurrenceRule is set.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_APPOINTMENT_RECUR);
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", baseInfo({ recurrenceRule: undefined }), "");
            expect(value).toEqual(Buffer.alloc(0));
        });

        it("Resolves PidLidTimeZoneStruct to the encoded timezone blob.", () => {
            const session = makeSession();
            const id = namedId(session, PSETID_APPOINTMENT, LID_TIME_ZONE_STRUCT);
            const info = baseInfo({ timezone: "UTC" });
            const value = calendarEventValueFor(session, id, PropertyType.PtypBinary, "calendarEvent:e1", info, "");
            expect(value).toEqual(encodeTimeZoneStruct("UTC", info.startDate));
        });

        it("Resolves PidLidReminderSet/ReminderDelta.", () => {
            const session = makeSession();
            const setId = namedId(session, PSETID_COMMON, LID_REMINDER_SET);
            const deltaId = namedId(session, PSETID_COMMON, LID_REMINDER_DELTA);
            const withReminder = baseInfo({ reminderMinutesBeforeStart: 15 });
            const withoutReminder = baseInfo({ reminderMinutesBeforeStart: undefined });

            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", withReminder, "")).toBe(true);
            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", withoutReminder, "")).toBe(false);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", withReminder, "")).toBe(15);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", withoutReminder, "")).toBe(0);
        });

        it("Falls back to a default for PidLidReminderSet/Delta under the wrong property set (PSETID_Appointment instead of PSETID_Common).", () => {
            const session = makeSession();
            const setId = namedId(session, PSETID_APPOINTMENT, LID_REMINDER_SET);
            const deltaId = namedId(session, PSETID_APPOINTMENT, LID_REMINDER_DELTA);
            expect(calendarEventValueFor(session, setId, PropertyType.PtypBoolean, "calendarEvent:e1", baseInfo(), "")).toBe(false);
            expect(calendarEventValueFor(session, deltaId, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "")).toBe(0);
        });

        it("Falls back to a default for every PSETID_Appointment-only LID under the wrong property set (PSETID_Common instead).", () => {
            const session = makeSession();
            const startId = namedId(session, PSETID_COMMON, LID_APPOINTMENT_START_WHOLE);
            const endId = namedId(session, PSETID_COMMON, LID_APPOINTMENT_END_WHOLE);
            const busyId = namedId(session, PSETID_COMMON, LID_BUSY_STATUS);
            const recurringId = namedId(session, PSETID_COMMON, LID_RECURRING);
            const timeZoneId = namedId(session, PSETID_COMMON, LID_TIME_ZONE_STRUCT);
            const responseId = namedId(session, PSETID_COMMON, LID_RESPONSE_STATUS);
            const info = baseInfo();

            expect(calendarEventValueFor(session, startId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toEqual(new Date(0));
            expect(calendarEventValueFor(session, endId, PropertyType.PtypTime, "calendarEvent:e1", info, "")).toEqual(new Date(0));
            expect(calendarEventValueFor(session, busyId, PropertyType.PtypInteger32, "calendarEvent:e1", info, "")).toBe(0);
            expect(calendarEventValueFor(session, recurringId, PropertyType.PtypBoolean, "calendarEvent:e1", info, "")).toBe(false);
            expect(calendarEventValueFor(session, timeZoneId, PropertyType.PtypBinary, "calendarEvent:e1", info, "")).toEqual(Buffer.alloc(0));
            expect(calendarEventValueFor(session, responseId, PropertyType.PtypInteger32, "calendarEvent:e1", info, "")).toBe(0);
        });

        describe("PidLidResponseStatus", () => {
            const responseId = (session: MapiSessionContext) => namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            it("Returns respOrganized (1) when the caller is the organizer.", () => {
                const session = makeSession();
                const id = responseId(session);
                const info = baseInfo({ organizerAddress: "Owner@Example.com" });
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", info, "owner@example.com")).toBe(1);
            });

            it("Returns the matching attendee's response status (by address, case-insensitive) when the caller is an attendee.", () => {
                const session = makeSession();
                const id = responseId(session);
                const attendees: Attendee[] = [
                    { address: "Attendee@Example.com", role: AttendeeRole.REQUIRED, responseStatus: AttendeeResponseStatus.ACCEPTED, isOrganizer: false },
                ];
                const info = baseInfo({ attendees });
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", info, "attendee@example.com")).toBe(3);
            });

            it("Returns respNone (0) when the caller is neither the organizer nor a listed attendee.", () => {
                const session = makeSession();
                const id = responseId(session);
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "nobody@example.com")).toBe(0);
            });

            it("Returns respNone (0) when callerAddress is empty (no mailbox address resolved).", () => {
                const session = makeSession();
                const id = responseId(session);
                expect(calendarEventValueFor(session, id, PropertyType.PtypInteger32, "calendarEvent:e1", baseInfo(), "")).toBe(0);
            });
        });
    });

    describe("resolvePropertyValues (calendarEvent dispatch)", () => {
        it("Resolves calendarEvent columns using the mailbox's own primarySmtpAddress as the caller address, without touching folderRepo/messageRepo.", async () => {
            const session = makeSession();
            const calendarEventRepo = {
                findOne: vi.fn().mockResolvedValue({
                    uid: "e1",
                    title: "Standup",
                    startDate: new Date("2026-09-07T14:00:00.000Z"),
                    endDate: new Date("2026-09-07T15:00:00.000Z"),
                    timezone: "UTC",
                    busyStatus: BusyStatus.BUSY,
                    organizer: { address: "organizer@example.com" },
                    attendees: [],
                }),
            };
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue({ primarySmtpAddress: "organizer@example.com" }) };
            const folderRepo = { findOne: vi.fn(), find: vi.fn() };
            const messageRepo = { findOne: vi.fn(), find: vi.fn() };
            const responseId = namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            const values = await resolvePropertyValues(
                "calendarEvent:e1",
                [
                    { propertyId: PID_TAG_SUBJECT, propertyType: PropertyType.PtypString },
                    { propertyId: responseId, propertyType: PropertyType.PtypInteger32 },
                ],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: folderRepo as any,
                    messageRepo: messageRepo as any,
                    calendarEventRepo: calendarEventRepo as any,
                    mailboxRepo: mailboxRepo as any,
                },
            );

            expect(values).toEqual(["Standup", 1]); // respOrganized, since the mailbox owns organizer@example.com
            expect(calendarEventRepo.findOne).toHaveBeenCalledWith("e1", { ignoreACL: true });
            expect(mailboxRepo.findOne).toHaveBeenCalledWith("mailbox-1", { ignoreACL: true });
            expect(folderRepo.findOne).not.toHaveBeenCalled();
            expect(messageRepo.findOne).not.toHaveBeenCalled();
        });

        it("Resolves callerAddress to empty string when the mailbox lookup itself comes back empty.", async () => {
            const session = makeSession();
            const calendarEventRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
            const mailboxRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
            const responseId = namedId(session, PSETID_APPOINTMENT, LID_RESPONSE_STATUS);

            const values = await resolvePropertyValues(
                "calendarEvent:gone",
                [{ propertyId: responseId, propertyType: PropertyType.PtypInteger32 }],
                {
                    mailboxUid: "mailbox-1",
                    session,
                    folderRepo: {} as any,
                    messageRepo: {} as any,
                    calendarEventRepo: calendarEventRepo as any,
                    mailboxRepo: mailboxRepo as any,
                },
            );

            expect(values).toEqual([0]); // respNone
        });
    });
});
