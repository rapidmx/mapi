///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { Message } from "@rapidmx/restapi";
import type { MapiSessionContext } from "../MapiSessionManager.js";

/**
 * The `RopGetContentsTable` analog of `FolderTarget.ts`: resolves a `"message:<uid>"` row target (see
 * `RopGetContentsTableHandler`'s own doc comment for why contents-table rows use this prefix instead of
 * `FolderTarget.ts`'s `"folder:"`/`"virtual:"`) into the display data a message-table row needs.
 */
export interface MessageTargetInfo {
    subject: string;
    read: boolean;
    hasAttachments: boolean;
    receivedDate: Date;
    /** `Label.uid`s applied to this message (Gmail-style, independent of folder placement) - resolved to
     * display names for `PidNameKeywords` (Outlook Categories) by `PropertyResolvers.ts`, which is where the
     * `labelRepo` needed to translate a uid into a name lives. */
    labelUids: string[];
}

/** Degrades to empty-looking values for a `"message:<uid>"` target whose real `Message` has since vanished
 * (soft-deleted or otherwise), the same "don't fail the whole ROP over one stale row" principle
 * `FolderTarget.resolveFolderInfo` already applies. */
export async function resolveMessageInfo(target: string, messageRepo: RepoUtils<any>): Promise<MessageTargetInfo> {
    const uid = target.slice("message:".length);
    const message: Message | undefined = await messageRepo.findOne(uid, { ignoreACL: true });
    return {
        subject: message?.subject ?? "",
        read: message?.flags?.read ?? false,
        hasAttachments: message?.hasAttachments ?? false,
        receivedDate: message?.receivedDate ?? new Date(0),
        labelUids: message?.labelUids ?? [],
    };
}

/** Resolves the messages directly in `folderUid`, as `"message:<uid>"` target strings, for
 * `RopGetContentsTable`. */
export async function resolveFolderMessages(folderUid: string, messageRepo: RepoUtils<any>): Promise<string[]> {
    const messages: Message[] = await messageRepo.find({ folderUid }, { ignoreACL: true });
    return messages.map((m) => `message:${m.uid}`);
}

/**
 * Returns `target`'s existing MID if an earlier `RopQueryRows` row already assigned one, otherwise assigns and
 * remembers the next free small integer MID. The exact `FolderTarget.assignOrGetFid` pattern, adapted for
 * messages: this is what lets a client `RopOpenMessage` a message it only ever learned about via a
 * `RopQueryRows` row's `PidTagMid` column.
 *
 * Backed by `session.messageTargetIds` (target -> MID) and `session.nextMessageId`, an O(1) reverse index/
 * counter pair rather than a linear scan of `session.messageIds` plus a `Math.max(...spread)` over its keys -
 * see `FolderTarget.assignOrGetFid`'s own doc comment for why both of those were real costs (not just
 * theoretical ones) at real mailbox/session scale: `messageIds` only ever grows for a session's lifetime as a
 * client pages through a mailbox, so a linear-scan lookup repeated once per row is quadratic over a session
 * that pages through many messages.
 */
export function assignOrGetMid(session: MapiSessionContext, target: string): number {
    const existing = session.messageTargetIds[target];
    if (existing !== undefined) {
        return existing;
    }
    const mid = session.nextMessageId++;
    session.messageIds[String(mid)] = target;
    session.messageTargetIds[target] = mid;
    return mid;
}
