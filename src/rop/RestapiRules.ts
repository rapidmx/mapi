///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { createHash } from "crypto";
import { BaseEntity, type RepoUtils } from "@rapidrest/service-core";

// Copies of restapi rules that @rapidmx/restapi 0.9.0 doesn't export from its package root. Keep them in sync with
// restapi's `util/ConversationUtils.ts` (`boundIndexedValue`) and `util/EntityUtils.ts` (`asEntity`).

/** The longest value `boundIndexedValue()` stores verbatim - restapi's `MAX_INDEXED_VALUE_LENGTH`. */
export const MAX_INDEXED_VALUE_LENGTH = 255;

/**
 * restapi's `boundIndexedValue`: `value` unchanged when it's at most `MAX_INDEXED_VALUE_LENGTH` characters, otherwise
 * `sha256:<64 hex digits>` of its UTF-8 bytes. restapi stores `CalendarEvent.icalUid` (and `Message.messageId`/
 * `conversationId`) this way, so a lookup by one of those identifiers has to bound its value the same way to match.
 */
export function boundIndexedValue<T extends string | null | undefined>(value: T): T {
    if (typeof value !== "string" || value.length <= MAX_INDEXED_VALUE_LENGTH) {
        return value;
    }
    return `sha256:${createHash("sha256").update(value, "utf8").digest("hex")}` as T;
}

/**
 * restapi's `asEntity`: `row` as an instance of `repo`'s model class. `RepoUtils.update()` only enforces its optimistic
 * lock when `existing instanceof BaseEntity`, and the Mongo backend's `find()`/`findOne()` return plain documents, so
 * passing one straight through turns a versioned update into an unconditional overwrite. SQL rows are already
 * instances and pass through as they are.
 */
export function asEntity<T>(repo: RepoUtils<any>, row: T): T {
    if (row instanceof BaseEntity) {
        return row;
    }
    const modelClass: any = (repo as any).modelClass;
    return modelClass ? new modelClass(row) : row;
}

/**
 * A string for a `RepoUtils` query that always matches `value` literally. A bare query string of the form `op(x)`
 * (`ne(x)`, `regex(x)`, ...) is read by service-core as an operator, so a client- or sender-supplied value is wrapped in
 * `eq(...)`, service-core's literal escape (its operator pattern is greedy, so `eq(ne(x))` compares against `ne(x)`).
 * service-core may still coerce the operand (`null`, numbers, `me`), so callers also compare the returned rows'
 * field against `value` themselves.
 */
export function literalQueryValue(value: string): string {
    return `eq(${value})`;
}
