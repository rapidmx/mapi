///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/** Splits a `PidTagDisplayTo`/`Cc`/`Bcc`-style string on the semicolons real Outlook separates recipients
 * with (also tolerating commas, in case a client or test harness uses that convention instead), trimming and
 * dropping empty entries. */
export function splitAddressList(value: string | undefined): string[] {
    if (!value) {
        return [];
    }
    return value
        .split(/[;,]/)
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

/** The plain addresses in a `PidTagDisplayTo`/`Cc`/`Bcc`-style string, skipping anything that isn't one (see
 * `isPlainEmailAddress`). */
export function parseAddressList(value: string | undefined): string[] {
    return splitAddressList(value).filter(isPlainEmailAddress);
}
