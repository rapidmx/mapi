///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { createHash } from "crypto";
import { BaseEntity, ModelUtils, type QueryLiteral, type RepoUtils } from "@rapidrest/service-core";

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
 * restapi's `asEntity`: `row` as an instance of `repo`'s model class. Before service-core 2.1.0, `RepoUtils.update()`
 * only enforced its optimistic lock when `existing instanceof BaseEntity`, and the Mongo backend's `find()`/`findOne()`
 * return plain documents, so passing one straight through turned a versioned update into an unconditional overwrite.
 * 2.1.0 also locks plain documents carrying a numeric `version`, so this is now defence in depth, kept while restapi
 * keeps its own copy. SQL rows are already instances and pass through as they are.
 */
export function asEntity<T>(repo: RepoUtils<any>, row: T): T {
    if (row instanceof BaseEntity) {
        return row;
    }
    const modelClass: any = (repo as any).modelClass;
    return modelClass ? new modelClass(row) : row;
}

/**
 * A `RepoUtils` query value that always matches `value` literally: service-core's `ModelUtils.literal()`, so a
 * client- or sender-supplied value like `ne(x)` is never read as an operator and `me`/`null`/numeric strings are not
 * substituted or coerced. Callers still compare the returned rows' field against `value` themselves.
 */
export function literalQueryValue(value: string): QueryLiteral {
    return ModelUtils.literal(value);
}
