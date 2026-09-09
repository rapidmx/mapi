///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/**
 * The Task named-property identity table (`(PropertySet GUID, LID)` pairs) this pragmatic subset supports, per
 * `[MS-OXPROPS]`/`[MS-OXOTASK]` - the exact `CalendarNamedProperties.ts` precedent, adapted for `PSETID_Task`,
 * shared by `PropertyResolvers.ts` (the read side: `taskValueFor`). Read-only in this pass - see
 * `TaskTarget.ts`'s own doc comment for why there's no write-side (`RopSetProperties`/`RopSaveChangesMessage`)
 * counterpart yet.
 */
export const PSETID_TASK = "00062003-0000-0000-c000-000000000046";

/** `PidLidTaskStatus` (`PT_LONG`) - `[MS-OXOTASK]` §2.2.2.1.5's own `olTaskNotStarted`(0)/`olTaskInProgress`(1)/
 * `olTaskComplete`(2)/`olTaskWaiting`(3)/`olTaskDeferred`(4) values. This data model's `Task.completed` is a
 * plain boolean with no in-progress/waiting/deferred distinction, so only `olTaskNotStarted`/`olTaskComplete`
 * are ever produced - a documented, harmless approximation, the same category as `CalendarNamedProperties.ts`'s
 * own `olWorkingElsewhere` gap. */
export const LID_TASK_STATUS = 0x8101;
export const TASK_STATUS_NOT_STARTED = 0;
export const TASK_STATUS_COMPLETE = 2;

/** `PidLidPercentComplete` (`PT_R8`, `[MS-OXOTASK]` §2.2.2.1.6) - `0.0`-`1.0`. No partial-progress tracking in
 * this data model either, so this is always `1.0` (complete) or `0.0` (not started), mirroring
 * `LID_TASK_STATUS`'s own two-value approximation. */
export const LID_PERCENT_COMPLETE = 0x8102;

/** `PidLidTaskDueDate` (`PT_SYSTIME`, `[MS-OXOTASK]` §2.2.2.2.4). */
export const LID_TASK_DUE_DATE = 0x8105;

/** `PidLidTaskComplete` (`PT_BOOLEAN`, `[MS-OXOTASK]` §2.2.2.1.9) - maps directly to `Task.completed`. */
export const LID_TASK_COMPLETE = 0x811c;
