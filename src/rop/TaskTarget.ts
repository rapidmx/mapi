///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import { Task } from "@rapidmx/restapi";
import type { RepoSort } from "./RepoPaging.js";

/**
 * The `RopGetContentsTable`/`RopOpenMessage` analog of `MessageTarget.ts`, for a `Folder` of type `TASKS`:
 * resolves a `"task:<uid>"` row/handle target into the display data a task-item row or an opened task's
 * properties need. A `Task` is addressed by MID exactly the same way a `Message`/`CalendarEvent`/`Contact` is -
 * `MessageTarget.assignOrGetMid`/`session.messageIds` are reused as-is, so this file adds no MID-registry code
 * of its own.
 *
 * Read-only in this pragmatic subset - a client checking off a task (`RopSetProperties`/`RopSaveChangesMessage`
 * against a `"task:"` handle) is a documented gap, same as Contacts (see `ContactTarget.ts`'s own note); the
 * REST API's own `PUT`/`PATCH` `/tasks/:id` remains the way to mutate a `Task` today.
 */
export interface TaskTargetInfo {
    title: string;
    completed: boolean;
    dueDate?: Date;
}

/** Degrades to empty-looking values for a `"task:<uid>"` target whose real `Task` has since vanished
 * (soft-deleted or otherwise) - the same "don't fail the whole ROP over one stale row" principle
 * `MessageTarget.resolveMessageInfo`/`ContactTarget.resolveContactInfo` already apply. */
export async function resolveTaskInfo(target: string, taskRepo: RepoUtils<any>): Promise<TaskTargetInfo> {
    const uid = target.slice("task:".length);
    const task: Task | undefined = await taskRepo.findOne(uid, { ignoreACL: true });
    return {
        title: task?.title ?? "",
        completed: task?.completed ?? false,
        dueDate: task?.dueDate,
    };
}

/** The order a `TASKS` folder's contents table lists tasks in (see `ContentsTable.ts`). */
export const TASK_SORT: RepoSort = { title: "ASC", uid: "ASC" };
