///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/** The most rows (items or folders) whose properties one `Execute` resolves, across all of its ROPs. */
export const MAX_ROWS_RESOLVED_PER_EXECUTE = 20000;

/** The most bytes one `Execute` builds or parses (FastTransfer streams and message bodies), across all of its ROPs. */
export const MAX_BYTES_BUILT_PER_EXECUTE = 64 * 1024 * 1024;

/** Thrown when an `Execute` has used up its `ExecuteBudget`. The ROP that hit it fails; see `RopDispatcher`. */
export class WorkBudgetExceededError extends Error {
    public constructor(what: string) {
        super(`Execute: the per-request ${what} budget is exhausted.`);
        this.name = "WorkBudgetExceededError";
    }
}

/**
 * The work one `Execute` may do. Limits on ROP count and per-ROP sizes alone still let one request chain many
 * expensive ROPs (a dozen 10000-row FastTransfer dumps, or the same large body parsed by several streams), so rows
 * resolved and bytes built are counted across the whole request. Also dedupes message body parses within the
 * request (`bodies`).
 */
export class ExecuteBudget {
    public rowsRemaining: number;
    public bytesRemaining: number;
    /** Message uid -> its decoded body, so opening the same body twice in one request parses it once. */
    public readonly bodies = new Map<string, Promise<Buffer>>();

    public constructor(maxRows: number = MAX_ROWS_RESOLVED_PER_EXECUTE, maxBytes: number = MAX_BYTES_BUILT_PER_EXECUTE) {
        this.rowsRemaining = maxRows;
        this.bytesRemaining = maxBytes;
    }

    /** Uses `count` rows of the budget, throwing `WorkBudgetExceededError` once it is exhausted. */
    public chargeRows(count = 1): void {
        this.rowsRemaining -= count;
        if (this.rowsRemaining < 0) {
            throw new WorkBudgetExceededError("row");
        }
    }

    /** Uses `count` bytes of the budget, throwing `WorkBudgetExceededError` once it is exhausted. */
    public chargeBytes(count: number): void {
        this.bytesRemaining -= count;
        if (this.bytesRemaining < 0) {
            throw new WorkBudgetExceededError("byte");
        }
    }
}
