///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { resolveContactInfo } from "../../src/rop/ContactTarget.js";
import { ContactAddressKind } from "@rapidmx/restapi";

describe("ContactTarget Tests", () => {
    describe("resolveContactInfo", () => {
        it("Resolves a real contact's fields.", async () => {
            const contactRepo = {
                findOne: vi.fn().mockResolvedValue({
                    uid: "c1",
                    displayName: "Jane Doe",
                    givenName: "Jane",
                    surname: "Doe",
                    emails: [{ address: "jane@example.com", type: ContactAddressKind.WORK }],
                    phones: [{ phoneNumber: "555-1234", type: ContactAddressKind.WORK }],
                    company: "Acme",
                    jobTitle: "Engineer",
                }),
            };

            const info = await resolveContactInfo("contact:c1", contactRepo as any);

            expect(info).toEqual({
                displayName: "Jane Doe",
                givenName: "Jane",
                surname: "Doe",
                email: "jane@example.com",
                businessPhone: "555-1234",
                companyName: "Acme",
                jobTitle: "Engineer",
            });
            expect(contactRepo.findOne).toHaveBeenCalledWith("c1", { ignoreACL: true });
        });

        it("Degrades to empty-looking values for a contact target whose real Contact has since vanished.", async () => {
            const contactRepo = { findOne: vi.fn().mockResolvedValue(undefined) };

            const info = await resolveContactInfo("contact:gone", contactRepo as any);

            expect(info.displayName).toBe("");
            expect(info.givenName).toBeUndefined();
            expect(info.surname).toBeUndefined();
            expect(info.email).toBeUndefined();
            expect(info.businessPhone).toBeUndefined();
            expect(info.companyName).toBeUndefined();
            expect(info.jobTitle).toBeUndefined();
        });

        it("Degrades emails/phones to undefined when the contact has none at all.", async () => {
            const contactRepo = { findOne: vi.fn().mockResolvedValue({ uid: "c1", displayName: "No Contact Info", emails: [], phones: [] }) };

            const info = await resolveContactInfo("contact:c1", contactRepo as any);

            expect(info.email).toBeUndefined();
            expect(info.businessPhone).toBeUndefined();
        });
    });
});
