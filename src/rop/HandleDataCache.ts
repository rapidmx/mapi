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

/** The cache key for the handle at `handleIndex`, as of its current generation. A generation is a random nonce
 * (see `assignHandle`), so a key is never reused, not even by a request whose session changes were never saved. */
export function handleDataKey(sessionUid: string, handleIndex: number, generation: string | undefined): string {
    return `${sessionUid}:${handleIndex}:${generation ?? ""}`;
}

/**
 * Where per-handle data that has to outlive one `Execute` lives: a built FastTransfer stream, and the chunks a write
 * stream has accumulated. With Redis (`RedisHandleDataStore`) every replica sees the same data, so a client can page
 * a stream out or finish writing a body through any pod. Without it (`MemoryHandleDataStore`) the data is only in
 * this process.
 */
export interface HandleDataStore {
    get(key: string): Promise<Buffer | undefined>;
    set(key: string, value: Buffer, ttlSeconds: number): Promise<void>;
    /** Stores `data` as the chunk starting at byte `offset` of the stream at `key`. Writing the same offset again
     * replaces the chunk, so a retried write never duplicates bytes. */
    putChunk(key: string, offset: number, data: Buffer, ttlSeconds: number): Promise<void>;
    /** The first `size` bytes of the stream at `key`, assembled from its chunks, or `undefined` when a chunk is
     * missing (expired or evicted) and the stream can't be reassembled exactly. */
    readChunks(key: string, size: number): Promise<Buffer | undefined>;
    delete(key: string): Promise<void>;
}

/** Walks `chunkAt(offset)` from offset 0 until `size` bytes are assembled. */
async function assembleChunks(size: number, chunkAt: (offset: number) => Promise<Buffer | undefined> | Buffer | undefined): Promise<Buffer | undefined> {
    const parts: Buffer[] = [];
    let offset = 0;
    while (offset < size) {
        const chunk = await chunkAt(offset);
        if (!chunk || chunk.length === 0) {
            return undefined;
        }
        parts.push(chunk);
        offset += chunk.length;
    }
    return offset === size ? Buffer.concat(parts) : undefined;
}

/** A `HandleDataStore` in this process only, on a `HandleDataCache` (so it is byte-bounded and can evict). */
export class MemoryHandleDataStore implements HandleDataStore {
    public constructor(private readonly cache: HandleDataCache = handleDataCache) {}

    public async get(key: string): Promise<Buffer | undefined> {
        return this.cache.get(key);
    }

    public async set(key: string, value: Buffer): Promise<void> {
        this.cache.set(key, value);
    }

    public async putChunk(key: string, offset: number, data: Buffer): Promise<void> {
        this.cache.set(`${key}#${offset}`, data);
    }

    public readChunks(key: string, size: number): Promise<Buffer | undefined> {
        return assembleChunks(size, (offset) => this.cache.get(`${key}#${offset}`));
    }

    public async delete(key: string): Promise<void> {
        this.cache.delete(key);
    }
}

/**
 * A `HandleDataStore` on a node-redis client, with this process's `HandleDataCache` in front for reads. Values are
 * stored base64-encoded (a plain string works with any client configuration); a write stream is a hash of
 * `offset -> chunk`.
 */
export class RedisHandleDataStore implements HandleDataStore {
    public constructor(
        private readonly client: any,
        private readonly local: HandleDataCache = handleDataCache,
    ) {}

    public async get(key: string): Promise<Buffer | undefined> {
        const cached = this.local.get(key);
        if (cached) {
            return cached;
        }
        const stored: string | null = await this.client.get(redisKey(key));
        if (stored == null) {
            return undefined;
        }
        const value = Buffer.from(stored, "base64");
        this.local.set(key, value);
        return value;
    }

    public async set(key: string, value: Buffer, ttlSeconds: number): Promise<void> {
        this.local.set(key, value);
        await this.client.setEx(redisKey(key), ttlSeconds, value.toString("base64"));
    }

    public async putChunk(key: string, offset: number, data: Buffer, ttlSeconds: number): Promise<void> {
        await this.client.hSet(redisKey(key), String(offset), data.toString("base64"));
        await this.client.expire(redisKey(key), ttlSeconds);
    }

    public async readChunks(key: string, size: number): Promise<Buffer | undefined> {
        const chunks: Record<string, string> = (await this.client.hGetAll(redisKey(key))) ?? {};
        return assembleChunks(size, (offset) => (chunks[String(offset)] !== undefined ? Buffer.from(chunks[String(offset)], "base64") : undefined));
    }

    public async delete(key: string): Promise<void> {
        this.local.delete(key);
        await this.client.del(redisKey(key));
    }
}

function redisKey(key: string): string {
    return `mapi.handle.${key}`;
}

/** The store used when a `RopContext` doesn't carry one (unit tests, or a route built without a session manager). */
export const defaultHandleDataStore: HandleDataStore = new MemoryHandleDataStore();
