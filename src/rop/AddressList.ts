///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import addressparser from "nodemailer/lib/addressparser/index.js";
import type { RopContext } from "./RopHandler.js";

/** Splits a `PidTagDisplayTo`/`Cc`/`Bcc`-style string on the semicolons Outlook separates recipients with,
 * trimming and dropping empty entries. Commas are not separators: a resolved recipient's display name commonly
 * contains one (`Doe, Jane <jane@example.com>`). */
export function splitAddressList(value: string | undefined): string[] {
    if (!value) {
        return [];
    }
    return value
        .split(";")
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
}

/** A bare `local@domain.tld` address: no whitespace (so no CR/LF), display name, angle brackets, quoting or
 * separators. */
const PLAIN_EMAIL_ADDRESS = /^[^\s@<>()[\]",;:\\]+@[^\s@<>()[\]",;:\\]+\.[^\s@<>()[\]",;:\\.]+$/;

/** `true` for a plain SMTP address (see `PLAIN_EMAIL_ADDRESS`) no longer than RFC 5321's 254 characters. These
 * strings reach MIME headers, the SMTP envelope and iCalendar `mailto:` lines, where a CR/LF would inject content. */
export function isPlainEmailAddress(value: string): boolean {
    return value.length <= 254 && PLAIN_EMAIL_ADDRESS.test(value);
}

/** One entry of a display string: an address with an optional display name, a display name alone (`address`
 * undefined, still to be resolved), or something that can't be used (`invalid`). */
export interface RecipientEntry {
    name: string;
    address?: string;
    invalid?: boolean;
}

/** Control characters (CR/LF included) are replaced in display names; nodemailer encodes the rest. */
function cleanName(name: string): string {
    return [...name].map((char) => (char.charCodeAt(0) < 0x20 || char.charCodeAt(0) === 0x7f ? " " : char)).join("").replace(/ +/g, " ").trim();
}

/**
 * Parses one `;`-separated entry with nodemailer's `addressparser`: `jane@example.com`, `Jane <jane@example.com>`,
 * `"Doe, Jane" <jane@example.com>` or a bare display name `Jane Doe`. An entry holding several addresses yields one
 * recipient each. Any address that isn't a plain SMTP address marks the entry invalid.
 */
export function parseRecipientEntry(entry: string): RecipientEntry[] {
    const parts = addressparser(entry, { flatten: true });
    const addressed = parts.filter((part) => part.address);
    if (addressed.length === 0) {
        const name = cleanName(entry);
        // Something address-shaped that addressparser couldn't take apart isn't a display name to look up.
        return [/[@<>]/.test(name) ? { name, invalid: true } : { name }];
    }
    const nameOnlyParts = parts.filter((part) => !part.address && part.name).map((part) => part.name);
    return addressed.map((part) => {
        const name = cleanName(addressed.length === 1 ? [...nameOnlyParts, part.name].filter(Boolean).join(", ") : part.name);
        const address = String(part.address);
        return isPlainEmailAddress(address) && !/[\r\n]/.test(entry) ? { name, address } : { name, invalid: true };
    });
}

/** Every entry of a `PidTagDisplayTo`/`Cc`/`Bcc`-style string - see `parseRecipientEntry`. */
export function parseRecipientList(value: string | undefined): RecipientEntry[] {
    return splitAddressList(value).flatMap(parseRecipientEntry);
}

/** The plain addresses in a `PidTagDisplayTo`/`Cc`/`Bcc`-style string, including those written as
 * `Name <address>`, skipping display names without an address and anything invalid. */
export function parseAddressList(value: string | undefined): string[] {
    return parseRecipientList(value)
        .filter((entry) => entry.address !== undefined)
        .map((entry) => entry.address!);
}

/** A recipient ready to send to. */
export interface ResolvedRecipient {
    name: string;
    address: string;
}

/** The outcome of `resolveRecipientList`: the usable recipients, and the entries that were display names with no
 * single matching contact (`unresolved`) or not usable at all (`invalid`). */
export interface RecipientResolution {
    recipients: ResolvedRecipient[];
    unresolved: string[];
    invalid: string[];
}

/** The most contacts looked at when resolving one display name - two is enough to tell "exactly one" apart. */
const MAX_NAME_MATCHES = 2;

/**
 * Resolves a display string into addresses. An entry carrying an address is used as is; a bare display name (Outlook
 * writes names, not addresses, into `PidTagDisplayTo` once it has resolved a recipient against an address book) is
 * looked up among the caller's own contacts by exact display name and used when exactly one contact matches, with
 * that contact's first valid email address.
 */
export async function resolveRecipientList(
    value: string | undefined,
    context: Pick<RopContext, "mailboxUid" | "contactRepo">,
): Promise<RecipientResolution> {
    const resolution: RecipientResolution = { recipients: [], unresolved: [], invalid: [] };
    for (const entry of parseRecipientList(value)) {
        if (entry.invalid) {
            resolution.invalid.push(entry.name);
        } else if (entry.address) {
            resolution.recipients.push({ name: entry.name, address: entry.address });
        } else {
            const contacts: { emails?: { address: string }[] }[] = context.contactRepo
                ? await context.contactRepo.find(
                      { mailboxUid: context.mailboxUid, displayName: entry.name, limit: MAX_NAME_MATCHES },
                      { ignoreACL: true, limit: MAX_NAME_MATCHES },
                  )
                : [];
            const address = contacts.length === 1 ? contacts[0].emails?.map((email) => email.address).find(isPlainEmailAddress) : undefined;
            if (address) {
                resolution.recipients.push({ name: entry.name, address });
            } else {
                resolution.unresolved.push(entry.name);
            }
        }
    }
    return resolution;
}
