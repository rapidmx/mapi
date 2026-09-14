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
} from "../src/MapiSessionManager.js";

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

    private setValue(key: string, value: string, ttl: number): void {
        this.values.set(key, value);
        this.ttls.set(key, ttl);
    }
}

