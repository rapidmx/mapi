///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { HttpRequest, HttpResponse, RepoUtils } from "@rapidrest/service-core";
import { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import { PropertyType, type PropertyTag, type PropertyValueData } from "../codec/PropertyValue.js";
import { defaultValueForType } from "../rop/PropertyResolvers.js";
import {
    BLANK_STAT,
    extractContentRestrictionSearchTerm,
    readLargePropertyTagArray,
    readStat,
    writeAddressBookPropertyRow,
    writeLargePropertyTagArray,
    writeStat,
} from "./NspiCodec.js";

/** `PidTagDisplayName`/`PidTagEmailAddress` - the small default column set this pragmatic subset returns when
 * a client's `GetMatches` request omits its own `Columns` field, mirroring EAS's own GAL result shape
 * (`SearchCommand.ts`'s `DisplayName`/`EmailAddress`). */
const PID_TAG_DISPLAY_NAME = 0x3001;
const PID_TAG_EMAIL_ADDRESS = 0x3003;
const DEFAULT_COLUMNS: PropertyTag[] = [
    { propertyId: PID_TAG_DISPLAY_NAME, propertyType: PropertyType.PtypString },
    { propertyId: PID_TAG_EMAIL_ADDRESS, propertyType: PropertyType.PtypString },
];

/** Escapes regex metacharacters so a client-supplied search string is matched literally by `RepoUtils`' own
 * `like(pattern)` query operator - a from-scratch reimplementation of EAS's `SearchCommand.ts`'s identical
 * helper (not a cross-import: EAS and MAPI are independent protocol verticals in this codebase). */
function escapeForLikeQuery(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** A small, non-cryptographic FNV-1a 32-bit hash of a `Contact.uid`, used as that contact's `MinimalEntryID`
 * in a `GetMatches` response. Deterministic per `uid` but not persisted or globally coordinated - acceptable
 * here since no other operation in this pragmatic subset ever looks a `MinimalEntryID` back up (`NspiGetProps`/
 * `NspiQueryRows`/... are out of scope), so nothing depends on it staying stable beyond a single response. */
function minimalEntryIdFor(uid: string): number {
    let hash = 0x811c9dc5;
    for (let i = 0; i < uid.length; i++) {
        hash ^= uid.charCodeAt(i);
        hash = Math.imul(hash, 0x01000193);
    }
    return hash >>> 0;
}

interface ContactRow {
    uid: string;
    displayName: string;
    emails: { address: string }[];
}

function valueForColumn(contact: ContactRow, column: PropertyTag): PropertyValueData {
    switch (column.propertyId) {
        case PID_TAG_DISPLAY_NAME:
            return contact.displayName;
        case PID_TAG_EMAIL_ADDRESS:
            return contact.emails[0]?.address ?? "";
        default:
            return defaultValueForType(column.propertyType);
    }
}

/** Finds every `Contact` in `mailboxUid` matching `searchTerm` (a case-insensitive substring match against
 * `displayName`/`givenName`/`surname`/`company`, queried per-field and merged in memory - the same two-backend
 * workaround `SearchCommand.ts` already established for its own identical `$or`-is-Mongo-only /
 * `ILike()`-needs-`%`-wrapping gaps) - or every `Contact` in the mailbox at all when `searchTerm` is
 * `undefined` (no filter was supplied, or (`extractContentRestrictionSearchTerm`'s own contract) the filter
 * was already known-good by the time this is called). */
async function findMatchingContacts(
    mailboxUid: string,
    contactRepo: RepoUtils<any>,
    searchTerm: string | undefined,
    likePattern: (escaped: string) => string,
): Promise<ContactRow[]> {
    if (!searchTerm) {
        return contactRepo.find({ mailboxUid }, { ignoreACL: true });
    }

    const pattern = likePattern(escapeForLikeQuery(searchTerm));
    const perField = await Promise.all(
        ["displayName", "givenName", "surname", "company"].map((field) =>
            contactRepo.find({ mailboxUid, [field]: `like(${pattern})` } as any, { ignoreACL: true }),
        ),
    );
    const byUid = new Map<string, ContactRow>();
    for (const contact of perField.flat() as ContactRow[]) {
        byUid.set(contact.uid, contact);
    }
    return Array.from(byUid.values()).sort((a, b) => a.displayName.localeCompare(b.displayName));
}

/**
 * `GetMatches` request type (`[MS-OXCMAPIHTTP]` §2.2.5.5, backed by `[MS-OXNSPI]`'s own `NspiGetMatches`
 * method): searches the GAL (this library's `Contact` records, the same source of truth EAS's own `Search`
 * command already uses) for entries matching the client's restriction, returning up to `RowCount` matches as
 * `AddressBookPropertyRow`s. This is the real operation behind a client's own address-book "search as you
 * type" UI - the closest real NSPI equivalent to this plan's own "minimal prefix/substring lookups" scope.
 *
 * `HasState`/`State` (an input `STAT` scoping the search and reporting position) and `HasMinimalIds`/
 * `MinimalIds` (an "Explicit Table" restricting the search to a caller-supplied candidate set) are decoded to
 * advance the reader correctly but not honored - this pragmatic subset always searches the caller's entire
 * Contact list fresh on every call (no real per-container/table state, see `NspiBindHandler.ts`'s own doc
 * comment for the same reasoning applied to `Bind`). `HasFilter`'s `Filter` is decoded via
 * `NspiCodec.extractContentRestrictionSearchTerm` - see that function's own doc comment for why an
 * unsupported restriction type throws rather than degrading. `HasPropertyName` (an alternate open-by-named-
 * property addressing mode) is decoded to advance the reader correctly but never itself used to select a
 * search field.
 *
 * The response's own `State` is always `BLANK_STAT` (with `TotalRecs` set to the real match count) - no real
 * per-container `STAT` is tracked, matching this pragmatic subset's stated scope.
 *
 * @author Jean-Philippe Steinmetz
 */
export async function handleNspiGetMatches(req: HttpRequest, res: HttpResponse, mailboxUid: string, contactRepo: RepoUtils<any>, likePattern: (escaped: string) => string): Promise<void> {
    const reader = new BufferReader(req.rawBody ?? Buffer.alloc(0));
    reader.readUInt32LE(); // Reserved
    if (reader.readUInt8()) {
        readStat(reader); // State - not honored, see function doc comment
    }
    if (reader.readUInt8()) {
        // MinimalIds (an Explicit Table) - not honored, see function doc comment
        const minimalIdCount = reader.readUInt32LE();
        for (let i = 0; i < minimalIdCount; i++) {
            reader.readUInt32LE();
        }
    }
    reader.readUInt32LE(); // InterfaceOptionFlags - reserved, always 0
    const searchTerm = reader.readUInt8() ? extractContentRestrictionSearchTerm(reader) : undefined;
    if (reader.readUInt8()) {
        reader.readBytes(16); // PropertyNameGuid - not honored, see function doc comment
        reader.readUInt32LE(); // PropertyNameId
    }
    const rowCount = reader.readUInt32LE();
    const columns = reader.readUInt8() ? readLargePropertyTagArray(reader) : DEFAULT_COLUMNS;
    // AuxiliaryBufferSize/AuxiliaryBuffer intentionally left unread - no auxiliary-payload support.

    const matches = await findMatchingContacts(mailboxUid, contactRepo, searchTerm, likePattern);
    const page = matches.slice(0, rowCount);

    const body = new BufferWriter();
    body.writeUInt32LE(0); // StatusCode - success
    body.writeUInt32LE(0); // ErrorCode - success
    body.writeUInt8(1); // HasState
    writeStat(body, { ...BLANK_STAT, totalRecs: matches.length });
    body.writeUInt8(page.length > 0 ? 1 : 0); // HasMinimalIds
    if (page.length > 0) {
        body.writeUInt32LE(page.length);
        for (const contact of page) {
            body.writeUInt32LE(minimalEntryIdFor(contact.uid));
        }
    }
    body.writeUInt8(page.length > 0 ? 1 : 0); // HasColumnsAndRows
    if (page.length > 0) {
        writeLargePropertyTagArray(body, columns);
        body.writeUInt32LE(page.length);
        for (const contact of page) {
            writeAddressBookPropertyRow(
                body,
                columns,
                columns.map((column) => valueForColumn(contact, column)),
            );
        }
    }
    body.writeUInt32LE(0); // AuxiliaryBufferSize
    res.status(200).send(body.toBuffer());
}
