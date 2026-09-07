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

/** Resolves a target string into the display data a folder-table row needs. A `"folder:<uid>"` target whose
 * `Folder` has since been deleted (soft-deleted or otherwise vanished) degrades to empty-looking values rather
 * than throwing - the row simply won't be interesting to a client, not a reason to fail the whole ROP. */
export async function resolveFolderInfo(
    mailboxUid: string,
    target: string,
    folderRepo: RepoUtils<any>,
): Promise<FolderTargetInfo> {
    if (target.startsWith("folder:")) {
        const uid = target.slice("folder:".length);
        const folder: Folder | undefined = await folderRepo.findOne(uid, { ignoreACL: true });
        const hasChildren = (await resolveFolderChildren(mailboxUid, target, folderRepo)).length > 0;
        return {
            displayName: folder?.name ?? "",
            unreadCount: folder?.unreadCount ?? 0,
            totalCount: folder?.totalCount ?? 0,
            hasChildren,
        };
    }
    const name = target.slice("virtual:".length);
    const hasChildren = (await resolveFolderChildren(mailboxUid, target, folderRepo)).length > 0;
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
 */
export async function resolveFolderChildren(
    mailboxUid: string,
    target: string,
    folderRepo: RepoUtils<any>,
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

    const allFolders: Folder[] = await folderRepo.find({ mailboxUid }, { ignoreACL: true });
    return allFolders
        .filter((f) => (isTopLevel ? f.parentFolderUid == null : f.parentFolderUid === parentFolderUid))
        .map((f) => `folder:${f.uid}`);
}

/**
 * Returns `target`'s existing FID if `RopLogon` or an earlier `RopGetHierarchyTable` row already assigned
 * one (a linear scan of `session.folderIds` - there is no reverse index, but a mailbox's folder count in this
 * pragmatic subset is small enough that this is never a real cost), otherwise assigns and remembers the next
 * free small integer FID. This is what lets a client `RopOpenFolder` a child folder it only ever learned
 * about via a `RopQueryRows` row's `PidTagFolderId` column - without this, only the 13 `RopLogon`-time
 * special folders could ever be opened.
 */
export function assignOrGetFid(session: MapiSessionContext, target: string): number {
    for (const [fid, existingTarget] of Object.entries(session.folderIds)) {
        if (existingTarget === target) {
            return Number(fid);
        }
    }
    const nextFid = Math.max(0, ...Object.keys(session.folderIds).map(Number)) + 1;
    session.folderIds[String(nextFid)] = target;
    return nextFid;
}
