///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { assignOrGetNamedPropertyId, resolveNamedProperty } from "../../src/rop/NamedPropertyRegistry.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";

describe("NamedPropertyRegistry Tests", () => {
    describe("assignOrGetNamedPropertyId", () => {
        it("Assigns the first named property starting at 0x8000.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
            expect(id).toBe(0x8000);
        });

        it("Assigns sequential IDs for distinct property names and reuses an existing assignment.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const first = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
            const second = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8209 });
            expect(second).toBe(first + 1);

            const reused = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
            expect(reused).toBe(first);
        });

        it("Distinguishes properties by GUID, not just LID.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const first = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
            const second = assignOrGetNamedPropertyId(session, { guid: "00062008-0000-0000-c000-000000000046", kind: "lid", lid: 0x8208 });
            expect(second).not.toBe(first);
        });

        it("Distinguishes a Kind=name property from a Kind=lid property with a coincidentally equal-looking value.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const lidId = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 1 });
            const nameId = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "name", name: "1" });
            expect(nameId).not.toBe(lidId);
        });

        it("Is case-insensitive on the GUID.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const lower = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT.toLowerCase(), kind: "lid", lid: 1 });
            const upper = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT.toUpperCase(), kind: "lid", lid: 1 });
            expect(upper).toBe(lower);
        });

        it("Returns 0x0000 for a new name once the session's entire 0x8000-0xFFFF ID space is exhausted, without throwing.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            session.nextNamedPropertyId = 0xffff;
            const last = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 1 });
            expect(last).toBe(0xffff);

            const overflowed = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 2 });
            expect(overflowed).toBe(0x0000);
        });

        it("Still returns a name's real ID after exhaustion if it was registered before the registry filled up.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const early = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 1 });
            session.nextNamedPropertyId = 0x10000;

            expect(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 1 })).toBe(early);
            expect(assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 2 })).toBe(0x0000);
        });
    });

    describe("resolveNamedProperty", () => {
        it("Resolves an assigned ID back to its original PropertyName.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            const id = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
            expect(resolveNamedProperty(session, id)).toEqual({ guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
        });

        it("Returns undefined for an ID that was never assigned.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            expect(resolveNamedProperty(session, 0x8000)).toBeUndefined();
        });

        it("Skips past non-matching entries to find the right one when the registry holds several.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 1 });
            const secondId = assignOrGetNamedPropertyId(session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 2 });
            expect(resolveNamedProperty(session, secondId)).toEqual({ guid: PSETID_APPOINTMENT, kind: "lid", lid: 2 });
        });
    });
});
