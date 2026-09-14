///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";

/** `RepoUtils.find()`'s own largest page. */
export const REPO_PAGE_SIZE = 1000;

/** The most rows gathered by `findAllCapped`. */
export const MAX_COLLECTION_ROWS = 10000;

/** A sort order in `RepoUtils` query syntax, e.g. `{ receivedDate: "DESC", uid: "ASC" }`. Always end with a
 * unique field so pages don't overlap or skip rows that tie on the leading fields. */
export type RepoSort = Record<string, "ASC" | "DESC">;

/**
 * Fetches one page. `limit`/`page`/`sort` go in both the query (what the SQL query builder reads) and the options
 * (what Mongo reads); without them each backend silently returns its own 100-row default in an unspecified order.
 */
export async function findPage<T>(repo: RepoUtils<any>, query: Record<string, any>, sort: RepoSort, page: number, pageSize: number = REPO_PAGE_SIZE): Promise<T[]> {
    return repo.find({ ...query, sort, limit: pageSize, page } as any, { ignoreACL: true, limit: pageSize, page });
}

/**
 * Fetches every row matching `query`, in `sort` order, up to `cap` rows. `truncated` is `true` when more rows may
 * exist past the cap.
 */
export async function findAllCapped<T>(
    repo: RepoUtils<any>,
    query: Record<string, any>,
    sort: RepoSort,
    cap: number = MAX_COLLECTION_ROWS,
): Promise<{ items: T[]; truncated: boolean }> {
    const items: T[] = [];
    for (let page = 0; ; page++) {
        const rows: T[] = await findPage<T>(repo, query, sort, page);
        items.push(...rows);
        if (rows.length < REPO_PAGE_SIZE) {
            return { items: items.slice(0, cap), truncated: items.length > cap };
        }
        if (items.length >= cap) {
            return { items: items.slice(0, cap), truncated: true };
        }
    }
}

/**
 * Fetches the rows at positions `[start, start + count)` of `query` in `sort` order, reading only the pages that
 * window touches.
 */
export async function findWindow<T>(repo: RepoUtils<any>, query: Record<string, any>, sort: RepoSort, start: number, count: number): Promise<T[]> {
    const result: T[] = [];
    let position = start;
    while (result.length < count) {
        const page = Math.floor(position / REPO_PAGE_SIZE);
        const rows: T[] = await findPage<T>(repo, query, sort, page);
        const taken = rows.slice(position - page * REPO_PAGE_SIZE, position - page * REPO_PAGE_SIZE + (count - result.length));
        result.push(...taken);
        position += taken.length;
        if (rows.length < REPO_PAGE_SIZE) {
            break;
        }
    }
    return result;
}
