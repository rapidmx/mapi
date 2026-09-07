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
    };
}

/** Resolves the messages directly in `folderUid`, as `"message:<uid>"` target strings, for
 * `RopGetContentsTable`. */
export async function resolveFolderMessages(folderUid: string, messageRepo: RepoUtils<any>): Promise<string[]> {
    const messages: Message[] = await messageRepo.find({ folderUid }, { ignoreACL: true });
    return messages.map((m) => `message:${m.uid}`);
}

/**
 * Returns `target`'s existing MID if an earlier `RopQueryRows` row already assigned one (a linear scan of
 * `session.messageIds` - there is no reverse index, but a single table's row count in this pragmatic subset is
 * never large enough for this to be a real cost), otherwise assigns and remembers the next free small integer
 * MID. The exact `FolderTarget.assignOrGetFid` pattern, adapted for messages: this is what lets a client
 * `RopOpenMessage` a message it only ever learned about via a `RopQueryRows` row's `PidTagMid` column.
 */
export function assignOrGetMid(session: MapiSessionContext, target: string): number {
    for (const [mid, existingTarget] of Object.entries(session.messageIds)) {
        if (existingTarget === target) {
            return Number(mid);
        }
    }
    const nextMid = Math.max(0, ...Object.keys(session.messageIds).map(Number)) + 1;
    session.messageIds[String(nextMid)] = target;
    return nextMid;
}
