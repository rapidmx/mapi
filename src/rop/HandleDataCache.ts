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

/** The most chunks one write stream keeps. Each `RopWriteStream` stores one chunk, so this bounds the hash a stream
 * builds up (and the work `readChunks` does) however small the writes are. */
export const MAX_CHUNKS_PER_STREAM = 8192;

/** How often, at most, an owner index that is over its quota is recounted from its members. Between recounts a write is
 * checked against the index's running total alone, so a client hammering a full quota can't make every write walk
 * every member. */
export const OWNER_RECOUNT_INTERVAL_MS = 1000;

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
 * would take any owner past its `maxBytes`. Sizes are raw (unencoded) bytes. Every write costs O(1) per owner: each
 * index keeps a running total, and a chunked stream keeps its own size (the end of its furthest chunk), so neither the
 * stream's chunks nor the index's members are walked. A delete given the same `owners` takes the key's bytes off their
 * totals. Data that disappears otherwise (TTL, eviction, `deleteOwnedBy` of another index, a delete without owners)
 * keeps counting until a write that the total would refuse recounts the index from its live members, at most once per
 * `OWNER_RECOUNT_INTERVAL_MS` per index. Totals only ever over-count, never under-count.
 */
export interface HandleDataStore {
    get(key: string): Promise<Buffer | undefined>;
    /** Stores `value` at `key`. `false` when an owner's quota refused it. */
    set(key: string, value: Buffer, ttlSeconds: number, owners?: HandleDataOwner[]): Promise<boolean>;
    /** Stores `data` as the chunk starting at byte `offset` of the stream at `key`. Writing the same offset again
     * replaces the chunk, so a retried write never duplicates bytes or charges the quota twice. `false` when an owner's
     * quota refused it, or the stream already has `MAX_CHUNKS_PER_STREAM` chunks and this is a new one. */
    putChunk(key: string, offset: number, data: Buffer, ttlSeconds: number, owners?: HandleDataOwner[]): Promise<boolean>;
    /** The first `size` bytes of the stream at `key`, assembled from its chunks, or `undefined` when a chunk is
     * missing (expired or evicted) and the stream can't be reassembled exactly. */
    readChunks(key: string, size: number): Promise<Buffer | undefined>;
    /** Deletes the value or chunked stream at `key`, taking it off the totals of `owners`. */
    delete(key: string, owners?: HandleDataOwner[]): Promise<void>;
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

/** An owner index in memory: its members' sizes, their running total, and when it may next be recounted. */
interface MemoryOwnerIndex {
    members: Map<string, number>;
    total: number;
    recountAfter: number;
}

/** A chunked stream in memory: the offsets of its chunks and its size (the end of its furthest chunk). */
interface MemoryChunkedStream {
    offsets: Set<number>;
    size: number;
}

/** A `HandleDataStore` in this process only, on a `HandleDataCache` (so it is byte-bounded and can evict). */
export class MemoryHandleDataStore implements HandleDataStore {
    private readonly indexes = new Map<string, MemoryOwnerIndex>();
    private readonly streams = new Map<string, MemoryChunkedStream>();

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
        const stream: MemoryChunkedStream = this.streams.get(key) ?? { offsets: new Set<number>(), size: 0 };
        if (!stream.offsets.has(offset) && stream.offsets.size >= MAX_CHUNKS_PER_STREAM) {
            return false;
        }
        const size = Math.max(stream.size, offset + data.length);
        if (!this.admit(key, size, owners)) {
            return false;
        }
        this.cache.set(`${key}#${offset}`, data);
        stream.offsets.add(offset);
        stream.size = size;
        this.streams.set(key, stream);
        return true;
    }

    public readChunks(key: string, size: number): Promise<Buffer | undefined> {
        return assembleChunks(size, (offset) => this.cache.get(`${key}#${offset}`));
    }

    public async delete(key: string, owners: HandleDataOwner[] = []): Promise<void> {
        this.cache.delete(key);
        for (const offset of this.streams.get(key)?.offsets ?? []) {
            this.cache.delete(`${key}#${offset}`);
        }
        this.streams.delete(key);
        for (const owner of owners) {
            const index = this.indexes.get(owner.index);
            const bytes = index?.members.get(key);
            if (index && bytes !== undefined) {
                index.members.delete(key);
                index.total -= bytes;
            }
        }
    }

    public async deleteOwnedBy(index: string): Promise<void> {
        for (const key of this.indexes.get(index)?.members.keys() ?? []) {
            await this.delete(key);
        }
        this.indexes.delete(index);
    }

    /** `true` while `key` still holds a value or its stream's first chunk (a stream missing that can't be read). */
    private exists(key: string): boolean {
        if (this.cache.sizeOf(key) !== undefined) {
            return true;
        }
        const first: number | undefined = this.streams.get(key)?.offsets.values().next().value;
        return first !== undefined && this.cache.sizeOf(`${key}#${first}`) !== undefined;
    }

    /** Records `key` at `bytes` in every owner index, unless that would take one past its quota. */
    private admit(key: string, bytes: number, owners: HandleDataOwner[]): boolean {
        const indexes: MemoryOwnerIndex[] = [];
        for (const owner of owners) {
            const index: MemoryOwnerIndex = this.indexes.get(owner.index) ?? { members: new Map<string, number>(), total: 0, recountAfter: 0 };
            this.indexes.set(owner.index, index);
            if (!this.fits(index, key, bytes, owner.maxBytes)) {
                return false;
            }
            indexes.push(index);
        }
        for (const index of indexes) {
            index.total += bytes - (index.members.get(key) ?? 0);
            index.members.set(key, bytes);
        }
        return true;
    }

    /** `true` when `key` at `bytes` keeps `index` within `maxBytes`, recounting an index that looks full (rate limited). */
    private fits(index: MemoryOwnerIndex, key: string, bytes: number, maxBytes: number): boolean {
        if (index.total - (index.members.get(key) ?? 0) + bytes <= maxBytes) {
            return true;
        }
        const now = Date.now();
        if (now < index.recountAfter) {
            return false;
        }
        index.recountAfter = now + OWNER_RECOUNT_INTERVAL_MS;
        index.total = 0;
        for (const [member, size] of index.members) {
            if (this.exists(member)) {
                index.total += size;
            } else {
                index.members.delete(member);
                this.streams.delete(member);
            }
        }
        return index.total - (index.members.get(key) ?? 0) + bytes <= maxBytes;
    }
}

/**
 * The owner index helpers every quota script shares. An index is a sorted set (member -> raw bytes) with its running
 * total in `<index>.total` (always the sum of the set's scores) and a recount rate limit in `<index>.recount`.
 * - `recount` rebuilds the total from the members whose data key still exists, dropping the rest.
 * - `admits` checks the running total, recounting (at most once per `recountMs`) only when that total would refuse. An
 * index without a total yet (written before totals existed, or expired) is recounted first.
 * - `record` sets a member's bytes, adjusting the total by the difference.
 * The data keys `recount` looks at are built from the members, so they aren't declared in KEYS (fine on one Redis node,
 * not on a cluster).
 */
const OWNER_INDEX_LUA = `
local function recount(index, prefix, indexTtl)
  local total = 0
  local members = redis.call('ZRANGE', index, 0, -1, 'WITHSCORES')
  for j = 1, #members, 2 do
    if redis.call('EXISTS', prefix .. members[j]) == 0 then
      redis.call('ZREM', index, members[j])
    else
      total = total + tonumber(members[j + 1])
    end
  end
  redis.call('SET', index .. '.total', tostring(total), 'EX', indexTtl)
  return total
end
local function admits(index, member, bytes, maxBytes, prefix, recountMs, indexTtl)
  local stored = redis.call('GET', index .. '.total')
  local fresh = not stored
  local total = stored and tonumber(stored) or recount(index, prefix, indexTtl)
  local old = tonumber(redis.call('ZSCORE', index, member) or 0)
  if total - old + bytes <= maxBytes then return true end
  if fresh or not redis.call('SET', index .. '.recount', '1', 'PX', recountMs, 'NX') then return false end
  total = recount(index, prefix, indexTtl)
  old = tonumber(redis.call('ZSCORE', index, member) or 0)
  return total - old + bytes <= maxBytes
end
local function record(index, member, bytes, indexTtl)
  local old = tonumber(redis.call('ZSCORE', index, member) or 0)
  redis.call('ZADD', index, tostring(bytes), member)
  redis.call('INCRBY', index .. '.total', tostring(bytes - old))
  redis.call('EXPIRE', index, indexTtl)
  redis.call('EXPIRE', index .. '.total', indexTtl)
end
`;

/** Stores a value within its owners' quotas. KEYS: data key, then each owner index. ARGV: data key prefix, member name,
 * index ttl, recount interval ms, base64 value, ttl, raw bytes, then each owner's max bytes. Returns 1 stored, 0 refused. */
export const SET_WITH_QUOTA_SCRIPT = `${OWNER_INDEX_LUA}
local bytes = tonumber(ARGV[7])
for i = 2, #KEYS do
  if not admits(KEYS[i], ARGV[2], bytes, tonumber(ARGV[i + 6]), ARGV[1], ARGV[4], ARGV[3]) then return 0 end
end
redis.call('SETEX', KEYS[1], ARGV[6], ARGV[5])
for i = 2, #KEYS do record(KEYS[i], ARGV[2], bytes, ARGV[3]) end
return 1
`;

/** Stores one write-stream chunk within its owners' quotas, in O(1): the hash keeps its size (the end of its furthest
 * chunk) in the `size` field, and that size is what the owners are charged, so a chunk rewritten at the same offset
 * isn't charged twice. KEYS: hash key, then each owner index. ARGV: data key prefix, member name, index ttl, recount
 * interval ms, unpadded base64url chunk, ttl, raw chunk bytes, hash field (offset), max chunks, then each owner's max
 * bytes. Returns 1 stored, 0 refused (quota, or a new chunk past the max chunks). */
export const PUT_CHUNK_WITH_QUOTA_SCRIPT = `${OWNER_INDEX_LUA}
local field = ARGV[8]
if redis.call('HEXISTS', KEYS[1], field) == 0 and redis.call('HLEN', KEYS[1]) > tonumber(ARGV[9]) then return 0 end
local size = math.max(tonumber(redis.call('HGET', KEYS[1], 'size') or 0), tonumber(field) + tonumber(ARGV[7]))
for i = 2, #KEYS do
  if not admits(KEYS[i], ARGV[2], size, tonumber(ARGV[i + 8]), ARGV[1], ARGV[4], ARGV[3]) then return 0 end
end
redis.call('HSET', KEYS[1], field, ARGV[5], 'size', tostring(size))
redis.call('EXPIRE', KEYS[1], ARGV[6])
for i = 2, #KEYS do record(KEYS[i], ARGV[2], size, ARGV[3]) end
return 1
`;

/** Deletes one data key and takes it off its owners' totals. KEYS: data key, then each owner index. ARGV: member name. */
export const DELETE_WITH_OWNERS_SCRIPT = `
redis.call('DEL', KEYS[1])
for i = 2, #KEYS do
  local old = redis.call('ZSCORE', KEYS[i], ARGV[1])
  if old then
    redis.call('ZREM', KEYS[i], ARGV[1])
    if redis.call('EXISTS', KEYS[i] .. '.total') == 1 then redis.call('INCRBY', KEYS[i] .. '.total', tostring(-tonumber(old))) end
  end
end
return 1
`;

/** Deletes every key an owner index names, then the index with its total and recount marker. KEYS: index. ARGV: data key
 * prefix. Returns the members (unprefixed keys) it deleted, so the caller can drop its local copies too. */
export const DELETE_OWNED_SCRIPT = `
local members = redis.call('ZRANGE', KEYS[1], 0, -1)
for _, member in ipairs(members) do
  redis.call('DEL', ARGV[1] .. member)
end
redis.call('DEL', KEYS[1], KEYS[1] .. '.total', KEYS[1] .. '.recount')
return members
`;

/** The prefix of every handle data key in Redis. */
export const HANDLE_DATA_PREFIX = "mapi.handle.";

/**
 * A `HandleDataStore` on a node-redis client, with this process's `HandleDataCache` in front for reads. Values are
 * stored base64-encoded (a plain string works with any client configuration); a write stream is a hash of
 * `offset -> chunk` plus its `size`. Owner indexes are sorted sets (`mapi.handle-owner.<index>`) of key -> raw bytes, with
 * a running total beside each (see `OWNER_INDEX_LUA`).
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
                ...ownerScriptArguments(key),
                value.toString("base64"),
                String(ttlSeconds),
                String(value.length),
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
                ...ownerScriptArguments(key),
                // Unpadded base64url: the smallest encoding Node's "base64" decoding reads back.
                data.toString("base64url"),
                String(ttlSeconds),
                String(data.length),
                String(offset),
                String(MAX_CHUNKS_PER_STREAM),
                ...owners.map((owner) => String(owner.maxBytes)),
            ],
        });
        return Number(result) === 1;
    }

    public async readChunks(key: string, size: number): Promise<Buffer | undefined> {
        const chunks: Record<string, string> = (await this.client.hGetAll(redisKey(key))) ?? {};
        return assembleChunks(size, (offset) => (chunks[String(offset)] !== undefined ? Buffer.from(chunks[String(offset)], "base64") : undefined));
    }

    public async delete(key: string, owners: HandleDataOwner[] = []): Promise<void> {
        this.local.delete(key);
        if (owners.length === 0) {
            await this.client.del(redisKey(key));
            return;
        }
        await this.client.eval(DELETE_WITH_OWNERS_SCRIPT, { keys: [redisKey(key), ...owners.map((owner) => ownerKey(owner.index))], arguments: [key] });
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

/** The ARGV every quota script starts with: data key prefix, member name, index ttl, recount interval ms. */
function ownerScriptArguments(key: string): string[] {
    return [HANDLE_DATA_PREFIX, key, String(HANDLE_DATA_INDEX_TTL_SECONDS), String(OWNER_RECOUNT_INTERVAL_MS)];
}

/** The Redis key of an owner index. */
export function ownerKey(index: string): string {
    return `mapi.handle-owner.${index}`;
}

/** The store used when a `RopContext` doesn't carry one (unit tests, or a route built without a session manager). */
export const defaultHandleDataStore: HandleDataStore = new MemoryHandleDataStore();
