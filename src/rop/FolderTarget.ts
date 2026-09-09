///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { Folder } from "@rapidmx/restapi";
import type { MapiSessionContext } from "../MapiSessionManager.js";

/**
 * Shared helpers for resolving the `"virtual:<name>"`/`"folder:<uid>"` target strings `RopLogonHandler`
 * assigns to FIDs (see its own doc comment) into real folder data - used by `RopOpenFolderHandler` and
 * `RopGetHierarchyTableHandler`.
 */

/** Display names for the virtual (no real backing `Folder`) special folders - see `RopLogonHandler`'s own
 * doc comment for why these don't have real rows in this data model. */
const VIRTUAL_DISPLAY_NAMES: Record<string, string> = {
    root: "Root",
    deferredAction: "Deferred Action",
    spoolerQueue: "Spooler Queue",
    ipmSubtree: "Top of Information Store",
    commonViews: "Common Views",
    schedule: "Schedule",
    search: "Finder",
    views: "Views",
    shortcuts: "Shortcuts",
};

export interface FolderTargetInfo {
    displayName: string;
    unreadCount: number;
    totalCount: number;
    hasChildren: boolean;
}

/** An optional, purely request-scoped (never persisted to the session) memo for `resolveFolderChildren`'s own
 * full-mailbox `folderRepo.find({ mailboxUid })` fetch - see that function's own doc comment for why this
 * matters. Deliberately just a plain object a caller creates fresh per ROP-handler invocation and threads
 * through every `resolveFolderInfo`/`resolvePropertyValues` call in that same batch, not session state - unlike
 * the FID/MID/named-property registries, there's no cross-request reuse to gain here (a mailbox's folder list
 * can change between requests), only cross-*row* reuse within one table page. */
export interface FolderResolutionCache {
    allFolders?: Folder[];
}

/** Resolves a target string into the display data a folder-table row needs. A `"folder:<uid>"` target whose
 * `Folder` has since been deleted (soft-deleted or otherwise vanished) degrades to empty-looking values rather
 * than throwing - the row simply won't be interesting to a client, not a reason to fail the whole ROP. */
export async function resolveFolderInfo(
    mailboxUid: string,
    target: string,
    folderRepo: RepoUtils<any>,
    cache?: FolderResolutionCache,
): Promise<FolderTargetInfo> {
    if (target.startsWith("folder:")) {
        const uid = target.slice("folder:".length);
        const folder: Folder | undefined = await folderRepo.findOne(uid, { ignoreACL: true });
        const hasChildren = (await resolveFolderChildren(mailboxUid, target, folderRepo, cache)).length > 0;
        return {
            displayName: folder?.name ?? "",
            unreadCount: folder?.unreadCount ?? 0,
            totalCount: folder?.totalCount ?? 0,
            hasChildren,
        };
    }
    const name = target.slice("virtual:".length);
    const hasChildren = (await resolveFolderChildren(mailboxUid, target, folderRepo, cache)).length > 0;
    return { displayName: VIRTUAL_DISPLAY_NAMES[name] ?? name, unreadCount: 0, totalCount: 0, hasChildren };
}

/**
 * Resolves the direct children of `target`, as an array of the same target-string format, for
 * `RopGetHierarchyTable`. Only `"virtual:root"`/`"virtual:ipmSubtree"` (this mailbox's top-level real
 * folders - both collapse to the same "top of the visible tree" concept in this pragmatic subset, see
 * `RopLogonHandler`'s own doc comment) and a real folder (its own real children) have any children at all;
 * every other virtual folder (Deferred Action, Spooler Queue, ...) is permanently empty, since this data
 * model has no concept of nesting anything under them.
 *
 * Filters an already-fetched full folder list in application code rather than querying by
 * `parentFolderUid: undefined` directly - query-DSL semantics for "field is unset" aren't reliably consistent
 * across backends (the exact kind of gap this project's own testing philosophy has caught before, e.g. `$or`
 * being Mongo-only), so filtering a fetched array sidesteps the question entirely rather than risking it.
 * The comparison itself uses `== null` (matching both `null` and `undefined`), not `=== undefined` - a
 * top-level folder's unset `parentFolderUid` round-trips as genuine `undefined` from Mongo but as `null` from
 * a SQL `nullable` column, a real, previously-confirmed cross-backend discrepancy in this codebase (caught by
 * this exact test against a real SQLite-backed server, not assumed).
 *
 * **`cache`**: without it, this fetches every folder in the mailbox on every call - fine for the single call
 * `RopGetHierarchyTableHandler` makes to build a table, but `resolveFolderInfo` also calls this once per row
 * just to compute `hasChildren`, and a table page can hold many rows. A caller resolving a whole page (
 * `RopQueryRowsHandler`/`RopGetPropertiesSpecificHandler`) passes one shared cache object through every call in
 * that batch so the full list is fetched at most once per batch instead of once per row - see
 * `FolderResolutionCache`'s own doc comment for why this isn't session state.
 */
export async function resolveFolderChildren(
    mailboxUid: string,
    target: string,
    folderRepo: RepoUtils<any>,
    cache?: FolderResolutionCache,
): Promise<string[]> {
    const isTopLevel = target === "virtual:root" || target === "virtual:ipmSubtree";
    let parentFolderUid: string | undefined;
    if (isTopLevel) {
        parentFolderUid = undefined;
    } else if (target.startsWith("folder:")) {
        parentFolderUid = target.slice("folder:".length);
    } else {
        return [];
    }

    let allFolders: Folder[];
    if (cache) {
        cache.allFolders ??= await folderRepo.find({ mailboxUid }, { ignoreACL: true });
        allFolders = cache.allFolders;
    } else {
        allFolders = await folderRepo.find({ mailboxUid }, { ignoreACL: true });
    }
    return allFolders
        .filter((f) => (isTopLevel ? f.parentFolderUid == null : f.parentFolderUid === parentFolderUid))
        .map((f) => `folder:${f.uid}`);
}

/**
 * Returns `target`'s existing FID if `RopLogon` or an earlier `RopGetHierarchyTable` row already assigned
 * one, otherwise assigns and remembers the next free small integer FID. This is what lets a client
 * `RopOpenFolder` a child folder it only ever learned about via a `RopQueryRows` row's `PidTagFolderId`
 * column - without this, only the 13 `RopLogon`-time special folders could ever be opened.
 *
 * Backed by `session.folderTargetIds` (target -> FID) and `session.nextFolderId`, an O(1) reverse index/counter
 * pair rather than a linear scan of `session.folderIds` plus a `Math.max(...spread)` over its keys - both real
 * costs, not just theoretical ones: `folderTargetIds` only ever grows for a session's lifetime as a client
 * browses more of a mailbox's folder tree, so the old scan-per-lookup approach was quadratic across a session
 * that opens many folders, and `Math.max` spreading an unbounded array as call arguments risked a stack
 * overflow past V8's argument-count limit on a mailbox with enough folders. `RopLogonHandler` keeps
 * `folderTargetIds`/`nextFolderId` in sync with `folderIds` for the 13 special folders it assigns directly,
 * rather than going through this function, since it assigns all 13 as one batch up front.
 */
export function assignOrGetFid(session: MapiSessionContext, target: string): number {
    const existing = session.folderTargetIds[target];
    if (existing !== undefined) {
        return existing;
    }
    const fid = session.nextFolderId++;
    session.folderIds[String(fid)] = target;
    session.folderTargetIds[target] = fid;
    return fid;
}
