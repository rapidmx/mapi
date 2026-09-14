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

    /** The size of the live entry at `key`, without refreshing it or changing its recency. */
    public sizeOf(key: string): number | undefined {
        const entry = this.entries.get(key);
        return entry && entry.expiresAt > Date.now() ? entry.value.length : undefined;
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

/** The `HandleDataStore` key of the write stream at `handleIndex`. */
export function writeStreamKey(sessionUid: string, handleIndex: number, generation: string | undefined): string {
    return `${handleDataKey(sessionUid, handleIndex, generation)}:write`;
}

/** The most bytes of FastTransfer streams and write-stream chunks one session keeps stored at once. */
export const MAX_HANDLE_DATA_BYTES_PER_SESSION = 64 * 1024 * 1024;

/** The most bytes of FastTransfer streams and write-stream chunks one user keeps stored at once, across all of their
 * sessions. */
export const MAX_HANDLE_DATA_BYTES_PER_USER = 128 * 1024 * 1024;

/** How long an owner index outlives its last write: a session's longest lifetime. */
export const HANDLE_DATA_INDEX_TTL_SECONDS = 24 * 60 * 60;

/** An index that stored data is counted against: the keys it holds and the most bytes they may add up to. */
export interface HandleDataOwner {
    index: string;
    maxBytes: number;
}

/** The index of everything a session stores, used for its quota and to delete it all when the session ends. */
export function sessionHandleDataIndex(sessionUid: string): string {
    return `session.${sessionUid}`;
}

/** The quota owners a write by `sessionUid` (belonging to `userUid`) counts against. */
export function handleDataOwners(sessionUid: string, userUid: string): HandleDataOwner[] {
    return [
        { index: sessionHandleDataIndex(sessionUid), maxBytes: MAX_HANDLE_DATA_BYTES_PER_SESSION },
        { index: `user.${userUid}`, maxBytes: MAX_HANDLE_DATA_BYTES_PER_USER },
    ];
}

/**
 * Where per-handle data that has to outlive one `Execute` lives: a built FastTransfer stream, and the chunks a write
 * stream has accumulated. With Redis (`RedisHandleDataStore`) every replica sees the same data, so a client can page
 * a stream out or finish writing a body through any pod. Without it (`MemoryHandleDataStore`) the data is only in
 * this process.
 *
 * **Quotas.** A write given `owners` is recorded in each owner's index and refused (`false`, nothing stored) when it
 * would take any owner past its `maxBytes`. Entries that have expired or been deleted stop counting the next time an
 * owner is checked. Sizes are raw (unencoded) bytes.
 */
export interface HandleDataStore {
    get(key: string): Promise<Buffer | undefined>;
    /** Stores `value` at `key`. `false` when an owner's quota refused it. */
    set(key: string, value: Buffer, ttlSeconds: number, owners?: HandleDataOwner[]): Promise<boolean>;
    /** Stores `data` as the chunk starting at byte `offset` of the stream at `key`. Writing the same offset again
     * replaces the chunk, so a retried write never duplicates bytes. `false` when an owner's quota refused it. */
    putChunk(key: string, offset: number, data: Buffer, ttlSeconds: number, owners?: HandleDataOwner[]): Promise<boolean>;
    /** The first `size` bytes of the stream at `key`, assembled from its chunks, or `undefined` when a chunk is
     * missing (expired or evicted) and the stream can't be reassembled exactly. */
    readChunks(key: string, size: number): Promise<Buffer | undefined>;
    /** Deletes the value or chunked stream at `key`. */
    delete(key: string): Promise<void>;
    /** Deletes everything recorded in the owner index `index`, and the index itself. */
    deleteOwnedBy(index: string): Promise<void>;
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
    /** Owner index -> key -> bytes. */
    private readonly indexes = new Map<string, Map<string, number>>();
    /** Chunked stream key -> the offsets of its chunks. */
    private readonly chunkOffsets = new Map<string, Set<number>>();

    public constructor(private readonly cache: HandleDataCache = handleDataCache) {}

    public async get(key: string): Promise<Buffer | undefined> {
        return this.cache.get(key);
    }

    public async set(key: string, value: Buffer, _ttlSeconds: number, owners: HandleDataOwner[] = []): Promise<boolean> {
        if (!this.admit(key, value.length, owners)) {
            return false;
        }
        this.cache.set(key, value);
        return true;
    }

    public async putChunk(key: string, offset: number, data: Buffer, _ttlSeconds: number, owners: HandleDataOwner[] = []): Promise<boolean> {
        const offsets = this.chunkOffsets.get(key) ?? new Set<number>();
        let total = data.length;
        for (const other of offsets) {
            total += other === offset ? 0 : (this.cache.sizeOf(`${key}#${other}`) ?? 0);
        }
        if (!this.admit(key, total, owners)) {
            return false;
        }
        this.cache.set(`${key}#${offset}`, data);
        this.chunkOffsets.set(key, offsets.add(offset));
        return true;
    }

    public readChunks(key: string, size: number): Promise<Buffer | undefined> {
        return assembleChunks(size, (offset) => this.cache.get(`${key}#${offset}`));
    }

    public async delete(key: string): Promise<void> {
        this.cache.delete(key);
        for (const offset of this.chunkOffsets.get(key) ?? []) {
            this.cache.delete(`${key}#${offset}`);
        }
        this.chunkOffsets.delete(key);
    }

    public async deleteOwnedBy(index: string): Promise<void> {
        for (const key of this.indexes.get(index)?.keys() ?? []) {
            await this.delete(key);
        }
        this.indexes.delete(index);
    }

    /** `true` while `key` still holds a value or at least one chunk. */
    private exists(key: string): boolean {
        if (this.cache.sizeOf(key) !== undefined) {
            return true;
        }
        return [...(this.chunkOffsets.get(key) ?? [])].some((offset) => this.cache.sizeOf(`${key}#${offset}`) !== undefined);
    }

    /** Records `key` at `bytes` in every owner index, unless that would take one past its quota. */
    private admit(key: string, bytes: number, owners: HandleDataOwner[]): boolean {
        for (const owner of owners) {
            const index = this.indexes.get(owner.index) ?? new Map<string, number>();
            let total = bytes;
            for (const [member, size] of index) {
                if (!this.exists(member)) {
                    index.delete(member);
                } else if (member !== key) {
                    total += size;
                }
            }
            if (total > owner.maxBytes) {
                return false;
            }
        }
        for (const owner of owners) {
            this.indexes.set(owner.index, (this.indexes.get(owner.index) ?? new Map<string, number>()).set(key, bytes));
        }
        return true;
    }
}

/** Sums an owner index's live members, dropping members whose key is gone and leaving out the member being written.
 * Shared by the two quota scripts: ARGV[5] is the data key prefix, ARGV[6] the member being written. */
const OWNER_TOTAL_LUA = `
local function ownerTotal(index)
  local total = 0
  local members = redis.call('ZRANGE', index, 0, -1, 'WITHSCORES')
  for j = 1, #members, 2 do
    if redis.call('EXISTS', ARGV[5] .. members[j]) == 0 then
      redis.call('ZREM', index, members[j])
    elseif members[j] ~= ARGV[6] then
      total = total + tonumber(members[j + 1])
    end
  end
  return total
end
`;

/** Stores a value within its owners' quotas. KEYS: data key, then each owner index. ARGV: base64 value, ttl, raw bytes,
 * index ttl, data key prefix, member name, then each owner's max bytes. Returns 1 stored, 0 refused. The data keys the
 * owner check looks at are built from the members, so they aren't declared in KEYS (fine on one Redis node, not on a
 * cluster). */
export const SET_WITH_QUOTA_SCRIPT = `${OWNER_TOTAL_LUA}
for i = 2, #KEYS do
  if ownerTotal(KEYS[i]) + tonumber(ARGV[3]) > tonumber(ARGV[i + 5]) then return 0 end
end
redis.call('SETEX', KEYS[1], ARGV[2], ARGV[1])
for i = 2, #KEYS do
  redis.call('ZADD', KEYS[i], ARGV[3], ARGV[6])
  redis.call('EXPIRE', KEYS[i], ARGV[4])
end
return 1
`;

/** Stores one write-stream chunk within its owners' quotas. KEYS: hash key, then each owner index. ARGV: unpadded
 * base64url chunk (so a stored chunk's raw size is exactly `floor(length * 3 / 4)`),
 * ttl, raw chunk bytes, index ttl, data key prefix, member name, hash field (offset), then each owner's max bytes. The
 * member counts as the sum of its chunks with this one replacing any chunk already at its offset. Returns 1 stored,
 * 0 refused. */
export const PUT_CHUNK_WITH_QUOTA_SCRIPT = `${OWNER_TOTAL_LUA}
local size = tonumber(ARGV[3])
for _, field in ipairs(redis.call('HKEYS', KEYS[1])) do
  if field ~= ARGV[7] then size = size + math.floor(redis.call('HSTRLEN', KEYS[1], field) * 3 / 4) end
end
for i = 2, #KEYS do
  if ownerTotal(KEYS[i]) + size > tonumber(ARGV[i + 6]) then return 0 end
end
redis.call('HSET', KEYS[1], ARGV[7], ARGV[1])
redis.call('EXPIRE', KEYS[1], ARGV[2])
for i = 2, #KEYS do
  redis.call('ZADD', KEYS[i], size, ARGV[6])
  redis.call('EXPIRE', KEYS[i], ARGV[4])
end
return 1
`;

/** Deletes every key an owner index names, then the index. KEYS: index. ARGV: data key prefix. Returns the members
 * (unprefixed keys) it deleted, so the caller can drop its local copies too. */
export const DELETE_OWNED_SCRIPT = `
local members = redis.call('ZRANGE', KEYS[1], 0, -1)
for _, member in ipairs(members) do
  redis.call('DEL', ARGV[1] .. member)
end
redis.call('DEL', KEYS[1])
return members
`;

/** The prefix of every handle data key in Redis. */
export const HANDLE_DATA_PREFIX = "mapi.handle.";

/**
 * A `HandleDataStore` on a node-redis client, with this process's `HandleDataCache` in front for reads. Values are
 * stored base64-encoded (a plain string works with any client configuration); a write stream is a hash of
 * `offset -> chunk`. Owner indexes are sorted sets (`mapi.handle-owner.<index>`) of key -> raw bytes.
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

    public async set(key: string, value: Buffer, ttlSeconds: number, owners: HandleDataOwner[] = []): Promise<boolean> {
        const result = await this.client.eval(SET_WITH_QUOTA_SCRIPT, {
            keys: [redisKey(key), ...owners.map((owner) => ownerKey(owner.index))],
            arguments: [
                value.toString("base64"),
                String(ttlSeconds),
                String(value.length),
                String(HANDLE_DATA_INDEX_TTL_SECONDS),
                HANDLE_DATA_PREFIX,
                key,
                ...owners.map((owner) => String(owner.maxBytes)),
            ],
        });
        if (Number(result) !== 1) {
            return false;
        }
        this.local.set(key, value);
        return true;
    }

    public async putChunk(key: string, offset: number, data: Buffer, ttlSeconds: number, owners: HandleDataOwner[] = []): Promise<boolean> {
        const result = await this.client.eval(PUT_CHUNK_WITH_QUOTA_SCRIPT, {
            keys: [redisKey(key), ...owners.map((owner) => ownerKey(owner.index))],
            arguments: [
                // Unpadded (base64url), so the script can tell a stored chunk's raw size from its length exactly. Node's
                // "base64" decoding reads it back.
                data.toString("base64url"),
                String(ttlSeconds),
                String(data.length),
                String(HANDLE_DATA_INDEX_TTL_SECONDS),
                HANDLE_DATA_PREFIX,
                key,
                String(offset),
                ...owners.map((owner) => String(owner.maxBytes)),
            ],
        });
        return Number(result) === 1;
    }

    public async readChunks(key: string, size: number): Promise<Buffer | undefined> {
        const chunks: Record<string, string> = (await this.client.hGetAll(redisKey(key))) ?? {};
        return assembleChunks(size, (offset) => (chunks[String(offset)] !== undefined ? Buffer.from(chunks[String(offset)], "base64") : undefined));
    }

    public async delete(key: string): Promise<void> {
        this.local.delete(key);
        await this.client.del(redisKey(key));
    }

    public async deleteOwnedBy(index: string): Promise<void> {
        const deleted: unknown[] = (await this.client.eval(DELETE_OWNED_SCRIPT, { keys: [ownerKey(index)], arguments: [HANDLE_DATA_PREFIX] })) ?? [];
        for (const key of deleted) {
            this.local.delete(String(key));
        }
    }
}

function redisKey(key: string): string {
    return `${HANDLE_DATA_PREFIX}${key}`;
}

/** The Redis key of an owner index. */
export function ownerKey(index: string): string {
    return `mapi.handle-owner.${index}`;
}

/** The store used when a `RopContext` doesn't carry one (unit tests, or a route built without a session manager). */
export const defaultHandleDataStore: HandleDataStore = new MemoryHandleDataStore();
