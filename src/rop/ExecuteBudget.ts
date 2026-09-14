///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/** The most rows (items or folders) whose properties one `Execute` resolves, across all of its ROPs. */
export const MAX_ROWS_RESOLVED_PER_EXECUTE = 20000;

/** The most bytes one `Execute` builds or parses (FastTransfer streams and message bodies), across all of its ROPs. */
export const MAX_BYTES_BUILT_PER_EXECUTE = 64 * 1024 * 1024;

/** The most database queries (paged reads, table lookups, deletes) one `Execute` runs through the budgeted paths. */
export const MAX_QUERIES_PER_EXECUTE = 20000;

/** The most rows one `Execute` reads back from paged queries (folder lists, folder walks, item lists). */
export const MAX_ROWS_FETCHED_PER_EXECUTE = 100000;

/** The most `RopSubmitMessage`s one `Execute` runs. Each one can relay mail, so a request can't repeat it 1024 times. */
export const MAX_SUBMITS_PER_EXECUTE = 16;

/** Thrown when an `Execute` has used up its `ExecuteBudget`. The ROP that hit it fails; see `RopDispatcher`. */
export class WorkBudgetExceededError extends Error {
    public constructor(what: string) {
        super(`Execute: the per-request ${what} budget is exhausted.`);
        this.name = "WorkBudgetExceededError";
    }
}

/**
 * The work one `Execute` may do. Limits on ROP count and per-ROP sizes alone still let one request chain many
 * expensive ROPs (a dozen 10000-row FastTransfer dumps, the same large body parsed by several streams, or a thousand
 * hierarchy tables or folder walks), so rows resolved, bytes built, queries run, rows fetched and submits are counted
 * across the whole request. Every charge happens before the work it pays for. Also dedupes message body parses within
 * the request (`bodies`).
 */
export class ExecuteBudget {
    public rowsRemaining: number;
    public bytesRemaining: number;
    public queriesRemaining: number;
    public fetchedRowsRemaining: number;
    public submitsRemaining: number;
    /** Message uid -> its decoded body, so opening the same body twice in one request parses it once. */
    public readonly bodies = new Map<string, Promise<Buffer>>();

    public constructor(
        maxRows: number = MAX_ROWS_RESOLVED_PER_EXECUTE,
        maxBytes: number = MAX_BYTES_BUILT_PER_EXECUTE,
        limits: { maxQueries?: number; maxFetchedRows?: number; maxSubmits?: number } = {},
    ) {
        this.rowsRemaining = maxRows;
        this.bytesRemaining = maxBytes;
        this.queriesRemaining = limits.maxQueries ?? MAX_QUERIES_PER_EXECUTE;
        this.fetchedRowsRemaining = limits.maxFetchedRows ?? MAX_ROWS_FETCHED_PER_EXECUTE;
        this.submitsRemaining = limits.maxSubmits ?? MAX_SUBMITS_PER_EXECUTE;
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

    /** Throws `WorkBudgetExceededError` when no bytes are left, so work that can't be sized up front (fetching a blob)
     * isn't started at all. */
    public assertBytesLeft(): void {
        if (this.bytesRemaining <= 0) {
            throw new WorkBudgetExceededError("byte");
        }
    }

    /** Uses `count` queries of the budget. Called before running them. */
    public chargeQueries(count = 1): void {
        this.queriesRemaining -= count;
        if (this.queriesRemaining < 0) {
            throw new WorkBudgetExceededError("query");
        }
    }

    /** Uses `count` fetched rows of the budget. */
    public chargeFetchedRows(count: number): void {
        this.fetchedRowsRemaining -= count;
        if (this.fetchedRowsRemaining < 0) {
            throw new WorkBudgetExceededError("fetched row");
        }
    }

    /** Uses one submit of the budget. */
    public chargeSubmit(): void {
        this.submitsRemaining -= 1;
        if (this.submitsRemaining < 0) {
            throw new WorkBudgetExceededError("submit");
        }
    }
}
