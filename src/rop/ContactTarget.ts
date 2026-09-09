///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { Contact } from "@rapidmx/restapi";

/**
 * The `RopGetContentsTable`/`RopOpenMessage` analog of `MessageTarget.ts`, for a `Folder` of type `CONTACTS`:
 * resolves a `"contact:<uid>"` row/handle target into the display data a contact-item row or an opened
 * address-card's properties need. A `Contact` is addressed by MID exactly the same way a `Message`/
 * `CalendarEvent` is - `MessageTarget.assignOrGetMid`/`session.messageIds` are reused as-is (see that file's own
 * doc comment), so this file adds no MID-registry code of its own.
 *
 * Read-only in this pragmatic subset - a client editing/creating a Contact via MAPI (`RopSetProperties`/
 * `RopSaveChangesMessage` against a `"contact:"` handle) is a documented gap, the same "browse only" limitation
 * this codebase already accepts for Tasks (`TaskTarget.ts`); NSPI's own `Bind`/`Unbind`/`GetMatches` already
 * covers the "search the GAL" case this data also backs.
 */
export interface ContactTargetInfo {
    displayName: string;
    givenName?: string;
    surname?: string;
    email?: string;
    businessPhone?: string;
    companyName?: string;
    jobTitle?: string;
}

/** Degrades to empty-looking values for a `"contact:<uid>"` target whose real `Contact` has since vanished
 * (soft-deleted or otherwise) - the same "don't fail the whole ROP over one stale row" principle
 * `MessageTarget.resolveMessageInfo`/`CalendarEventTarget.resolveCalendarEventInfo` already apply. */
export async function resolveContactInfo(target: string, contactRepo: RepoUtils<any>): Promise<ContactTargetInfo> {
    const uid = target.slice("contact:".length);
    const contact: Contact | undefined = await contactRepo.findOne(uid, { ignoreACL: true });
    return {
        displayName: contact?.displayName ?? "",
        givenName: contact?.givenName,
        surname: contact?.surname,
        email: contact?.emails?.[0]?.address,
        businessPhone: contact?.phones?.[0]?.phoneNumber,
        companyName: contact?.company,
        jobTitle: contact?.jobTitle,
    };
}

/** Resolves the contacts directly in `folderUid`, as `"contact:<uid>"` target strings, for
 * `RopGetContentsTable` against a `CONTACTS`-type folder. */
export async function resolveFolderContacts(folderUid: string, contactRepo: RepoUtils<any>): Promise<string[]> {
    const contacts: Contact[] = await contactRepo.find({ folderUid }, { ignoreACL: true });
    return contacts.map((c) => `contact:${c.uid}`);
}
