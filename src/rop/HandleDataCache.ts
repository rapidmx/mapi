///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/** How long an entry lives without being read - the same idle lifetime as a session. */
export const HANDLE_DATA_TTL_MS = 15 * 60 * 1000;

/** The most bytes this process keeps across all entries before evicting the least recently used ones. */
export const HANDLE_DATA_MAX_BYTES = 256 * 1024 * 1024;

interface Entry {
    value: Buffer;
    expiresAt: number;
}

/**
 * A process-local, byte-bounded LRU with TTL for large per-handle data: an opened stream's decoded message body
 * and a FastTransfer handle's built stream. Keeping these out of the session JSON means an `Execute` doesn't
 * re-serialize megabytes to Redis, and a `RopReadStream` doesn't re-fetch and re-parse the whole MIME message
 * for every chunk.
 *
 * It is only a cache. Every caller can rebuild an entry from the handle's own (small) session state, so a miss
 * (another replica served the previous request, eviction, TTL) costs one rebuild, never wrong data. Keys carry
 * the handle's generation (see `assignHandle`), so a reused handle index never picks up a previous handle's data.
 */
export class HandleDataCache {
    private readonly entries = new Map<string, Entry>();
    private totalBytes = 0;

    public constructor(
        private readonly maxBytes: number = HANDLE_DATA_MAX_BYTES,
        private readonly ttlMs: number = HANDLE_DATA_TTL_MS,
    ) {}

    public get(key: string): Buffer | undefined {
        const entry = this.entries.get(key);
        if (!entry) {
            return undefined;
        }
        this.entries.delete(key);
        if (entry.expiresAt <= Date.now()) {
            this.totalBytes -= entry.value.length;
            return undefined;
        }
        // Re-inserting moves the entry to the end of the Map's iteration order (most recently used).
        entry.expiresAt = Date.now() + this.ttlMs;
        this.entries.set(key, entry);
        return entry.value;
    }

    /** Stores `value` under `key`. A value larger than the whole budget isn't cached at all. */
    public set(key: string, value: Buffer): void {
        this.delete(key);
        if (value.length > this.maxBytes) {
            return;
        }
        this.entries.set(key, { value, expiresAt: Date.now() + this.ttlMs });
        this.totalBytes += value.length;
        for (const [oldestKey, oldest] of this.entries) {
            if (this.totalBytes <= this.maxBytes) {
                break;
            }
            this.entries.delete(oldestKey);
            this.totalBytes -= oldest.value.length;
        }
    }

    public delete(key: string): void {
        const entry = this.entries.get(key);
        if (entry) {
            this.entries.delete(key);
            this.totalBytes -= entry.value.length;
        }
    }

    /** The bytes currently held, for tests and diagnostics. */
    public get size(): number {
        return this.totalBytes;
    }
}

/** The process-wide instance every handler shares. */
export const handleDataCache = new HandleDataCache();

/** The cache key for the handle at `handleIndex`, as of its current generation. */
export function handleDataKey(sessionUid: string, handleIndex: number, generation: number | undefined): string {
    return `${sessionUid}:${handleIndex}:${generation ?? 0}`;
}
