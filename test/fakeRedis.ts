///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import {
    ACQUIRE_LOCK_SCRIPT,
    ADD_TO_USER_INDEX_SCRIPT,
    COMPARE_AND_SET_SCRIPT,
    CREATE_SCRIPT,
    RELEASE_LOCK_SCRIPT,
    RENEW_LOCK_SCRIPT,
    TOUCH_USER_INDEX_SCRIPT,
} from "../src/MapiSessionManager.js";
import { DELETE_OWNED_SCRIPT, DELETE_WITH_OWNERS_SCRIPT, PUT_CHUNK_WITH_QUOTA_SCRIPT, SET_WITH_QUOTA_SCRIPT } from "../src/rop/HandleDataCache.js";

/** A node-redis stand-in holding values in Maps. `eval` applies the semantics of each script the store runs, so the
 * store's handling of each result can be exercised without a Redis server. Scripts run synchronously, like Redis. */
export class FakeRedisClient {
    public readonly values = new Map<string, string>();
    public readonly hashes = new Map<string, Map<string, string>>();
    public readonly sortedSets = new Map<string, Map<string, number>>();
    public readonly ttls = new Map<string, number>();
    public readonly eval = vi.fn(async (script: string, options: { keys: string[]; arguments: string[] }) => {
        const keys = options.keys;
        const args = options.arguments;
        switch (script) {
            case CREATE_SCRIPT:
                this.setValue(keys[0], args[0], Number(args[1]));
                this.setValue(keys[1], "0", Number(args[1]));
                return 1;
            case COMPARE_AND_SET_SCRIPT: {
                const current = this.values.get(keys[1]);
                if (current === undefined || !this.values.has(keys[0])) {
                    return -1;
                }
                if (Number(current) !== Number(args[0])) {
                    return 0;
                }
                this.setValue(keys[0], args[1], Number(args[2]));
                this.setValue(keys[1], String(Number(args[0]) + 1), Number(args[2]));
                return 1;
            }
            case ADD_TO_USER_INDEX_SCRIPT: {
                const set = this.sortedSets.get(keys[0]) ?? new Map<string, number>();
                this.sortedSets.set(keys[0], set);
                const keyOf = (id: string): string => args[4] + id + args[5];
                for (const id of [...set.keys()]) {
                    if (!this.values.has(keyOf(id))) {
                        set.delete(id);
                    }
                }
                const ended: string[] = [];
                while (set.size >= Number(args[2])) {
                    const oldest = [...set.entries()].sort((a, b) => a[1] - b[1])[0][0];
                    set.delete(oldest);
                    this.values.delete(keyOf(oldest));
                    this.values.delete(`${keyOf(oldest)}.version`);
                    ended.push(oldest);
                }
                set.set(args[0], Number(args[1]));
                this.ttls.set(keys[0], Number(args[3]));
                return ended;
            }
            case ACQUIRE_LOCK_SCRIPT:
                if (this.values.has(keys[0])) {
                    return 0;
                }
                this.values.set(keys[0], args[0]);
                this.ttls.set(keys[0], Number(args[1]));
                return 1;
            case RELEASE_LOCK_SCRIPT:
                if (this.values.get(keys[0]) === args[0]) {
                    this.values.delete(keys[0]);
                }
                return 1;
            case RENEW_LOCK_SCRIPT:
                if (this.values.get(keys[0]) !== args[0]) {
                    return 0;
                }
                this.ttls.set(keys[0], Number(args[1]));
                return 1;
            case TOUCH_USER_INDEX_SCRIPT: {
                const set = this.sortedSets.get(keys[0]);
                if (set?.has(args[1])) {
                    set.set(args[1], Number(args[0]));
                }
                return 1;
            }
            case SET_WITH_QUOTA_SCRIPT: {
                const [prefix, member, indexTtl, recountMs] = args;
                const size = Number(args[6]);
                if (keys.slice(1).some((index, i) => !this.admits(index, member, size, Number(args[7 + i]), prefix, Number(recountMs)))) {
                    return 0;
                }
                this.setValue(keys[0], args[4], Number(args[5]));
                this.recordOwners(keys.slice(1), member, size, Number(indexTtl));
                return 1;
            }
            case PUT_CHUNK_WITH_QUOTA_SCRIPT: {
                const [prefix, member, indexTtl, recountMs] = args;
                const hash = this.hashes.get(keys[0]) ?? new Map<string, string>();
                const field = args[7];
                if (!hash.has(field) && hash.size > Number(args[8])) {
                    return 0;
                }
                const size = Math.max(Number(hash.get("size") ?? 0), Number(field) + Number(args[6]));
                if (keys.slice(1).some((index, i) => !this.admits(index, member, size, Number(args[9 + i]), prefix, Number(recountMs)))) {
                    return 0;
                }
                hash.set(field, args[4]);
                hash.set("size", String(size));
                this.hashes.set(keys[0], hash);
                this.ttls.set(keys[0], Number(args[5]));
                this.recordOwners(keys.slice(1), member, size, Number(indexTtl));
                return 1;
            }
            case DELETE_WITH_OWNERS_SCRIPT: {
                await this.del(keys[0]);
                for (const index of keys.slice(1)) {
                    const old = this.sortedSets.get(index)?.get(args[0]);
                    if (old !== undefined) {
                        this.sortedSets.get(index)!.delete(args[0]);
                        if (this.values.has(`${index}.total`)) {
                            this.values.set(`${index}.total`, String(Number(this.values.get(`${index}.total`)) - old));
                        }
                    }
                }
                return 1;
            }
            case DELETE_OWNED_SCRIPT: {
                const members = [...(this.sortedSets.get(keys[0])?.keys() ?? [])];
                for (const member of members) {
                    await this.del(args[0] + member);
                }
                this.sortedSets.delete(keys[0]);
                this.values.delete(`${keys[0]}.total`);
                this.recountMarkers.delete(`${keys[0]}.recount`);
                return members;
            }
            default:
                throw new Error("FakeRedisClient: unknown script");
        }
    });

    public async get(key: string): Promise<string | null> {
        return this.values.get(key) ?? null;
    }

    public async setEx(key: string, ttl: number, value: string): Promise<void> {
        this.setValue(key, value, ttl);
    }

    public async del(keys: string | string[]): Promise<void> {
        for (const key of Array.isArray(keys) ? keys : [keys]) {
            this.values.delete(key);
            this.hashes.delete(key);
        }
    }

    public async hSet(key: string, field: string, value: string): Promise<void> {
        const hash = this.hashes.get(key) ?? new Map<string, string>();
        hash.set(field, value);
        this.hashes.set(key, hash);
    }

    public async hGetAll(key: string): Promise<Record<string, string>> {
        return Object.fromEntries(this.hashes.get(key) ?? []);
    }

    public async expire(key: string, ttl: number): Promise<void> {
        this.ttls.set(key, ttl);
    }

    /** Counts of how many times each owner index was recounted from its members, for tests. */
    public readonly recounts = new Map<string, number>();
    /** Recount rate limit markers (`<index>.recount`) -> when they expire (ms). */
    public readonly recountMarkers = new Map<string, number>();

    /** Rebuilds an owner index's total from its members whose data still exists. */
    private recount(index: string, prefix: string): number {
        this.recounts.set(index, (this.recounts.get(index) ?? 0) + 1);
        const set = this.sortedSets.get(index) ?? new Map<string, number>();
        let total = 0;
        for (const [name, size] of set) {
            if (!this.values.has(prefix + name) && !this.hashes.has(prefix + name)) {
                set.delete(name);
            } else {
                total += size;
            }
        }
        this.values.set(`${index}.total`, String(total));
        return total;
    }

    /** The script's `admits`: the running total decides, recounted (rate limited) only when it would refuse. */
    private admits(index: string, member: string, bytes: number, maxBytes: number, prefix: string, recountMs: number): boolean {
        const stored = this.values.get(`${index}.total`);
        let total = stored !== undefined ? Number(stored) : this.recount(index, prefix);
        let old = this.sortedSets.get(index)?.get(member) ?? 0;
        if (total - old + bytes <= maxBytes) {
            return true;
        }
        const marker = `${index}.recount`;
        if (stored === undefined || (this.recountMarkers.get(marker) ?? 0) > Date.now()) {
            return false;
        }
        this.recountMarkers.set(marker, Date.now() + recountMs);
        total = this.recount(index, prefix);
        old = this.sortedSets.get(index)?.get(member) ?? 0;
        return total - old + bytes <= maxBytes;
    }

    private recordOwners(indexes: string[], member: string, size: number, ttl: number): void {
        for (const index of indexes) {
            const set = this.sortedSets.get(index) ?? new Map<string, number>();
            const old = set.get(member) ?? 0;
            this.sortedSets.set(index, set.set(member, size));
            this.values.set(`${index}.total`, String(Number(this.values.get(`${index}.total`) ?? 0) + size - old));
            this.ttls.set(index, ttl);
        }
    }

    private setValue(key: string, value: string, ttl: number): void {
        this.values.set(key, value);
        this.ttls.set(key, ttl);
    }
}

