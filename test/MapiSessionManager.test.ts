///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import {
    MapiSessionManager,
    MAX_SESSION_LIFETIME_MS,
    MAX_SESSIONS_PER_USER,
    MemoryMapiSessionStore,
    RedisMapiSessionStore,
    SESSION_TTL_SECONDS,
} from "../src/MapiSessionManager.js";

/** A node-redis stand-in holding values in a Map. `eval` applies the compare-and-set script's semantics (the only
 * script the store runs), so the store's handling of each result can be exercised without a Redis server. */
class FakeRedisClient {
    public readonly values = new Map<string, string>();
    public readonly ttls = new Map<string, number>();
    public readonly eval = vi.fn(async (_script: string, options: { keys: string[]; arguments: string[] }) => {
        const [key] = options.keys;
        const [expectedVersion, value, ttl] = options.arguments;
        const current = this.values.get(key);
        if (current === undefined) {
            return -1;
        }
        if (Number(JSON.parse(current).version) !== Number(expectedVersion)) {
            return 0;
        }
        this.values.set(key, value);
        this.ttls.set(key, Number(ttl));
        return 1;
    });

    public async get(key: string): Promise<string | null> {
        return this.values.get(key) ?? null;
    }

    public async exists(key: string): Promise<number> {
        return this.values.has(key) ? 1 : 0;
    }

    public async setEx(key: string, ttl: number, value: string): Promise<void> {
        this.values.set(key, value);
        this.ttls.set(key, ttl);
    }

    public async del(key: string): Promise<void> {
        this.values.delete(key);
    }
}

function memoryManager(): MapiSessionManager {
    const manager = new MapiSessionManager();
    manager.init();
    return manager;
}

function redisManager(client: FakeRedisClient): MapiSessionManager {
    const manager = new MapiSessionManager();
    (manager as any).redisClient = client;
    manager.init();
    return manager;
}

describe("MapiSessionManager Tests", () => {
    afterEach(() => {
        vi.useRealTimers();
    });

    describe.each([
        ["memory", () => ({ manager: memoryManager() })],
        ["redis", () => ({ manager: redisManager(new FakeRedisClient()) })],
    ])("on the %s store", (_name, build) => {
        it("Round-trips a session, handing out an independent copy on every load.", async () => {
            const { manager } = build();
            const created = await manager.create("mailbox-1", "user-1");

            const first = (await manager.load(created.uid))!;
            const second = (await manager.load(created.uid))!;
            first.handles[1] = { type: "logon", entityUid: "mailbox-1" };

            expect(first.mailboxUid).toBe("mailbox-1");
            expect(first.userUid).toBe("user-1");
            expect(second.handles[1]).toBeUndefined();
            expect(await manager.load("unknown")).toBeUndefined();
        });

        it("Saves compare-and-set on version: the second of two overlapping saves conflicts and keeps nothing.", async () => {
            const { manager } = build();
            const created = await manager.create("mailbox-1", "user-1");
            const requestA = (await manager.load(created.uid))!;
            const requestB = (await manager.load(created.uid))!;
            requestA.handles[1] = { type: "logon", entityUid: "from-a" };
            requestB.handles[1] = { type: "logon", entityUid: "from-b" };

            expect(await manager.save(requestA)).toBe("saved");
            expect(requestA.version).toBe(1);
            expect(await manager.save(requestB)).toBe("conflict");
            expect(requestB.version).toBe(0); // restored, so the object still describes what it was loaded as

            const stored = (await manager.load(created.uid))!;
            expect(stored.handles[1].entityUid).toBe("from-a");
            expect(stored.version).toBe(1);
            expect(await manager.save(stored)).toBe("saved");
        });

        it("Reports a save of a destroyed session as missing instead of recreating it.", async () => {
            const { manager } = build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            await manager.destroy(created.uid);

            expect(await manager.save(loaded)).toBe("missing");
            expect(await manager.load(created.uid)).toBeUndefined();
        });

        it("Ends a session older than MAX_SESSION_LIFETIME_MS on load, however recently it was used.", async () => {
            const { manager } = build();
            const created = await manager.create("mailbox-1", "user-1");
            const loaded = (await manager.load(created.uid))!;
            loaded.createdAt = new Date(Date.now() - MAX_SESSION_LIFETIME_MS - 1000).toISOString();
            expect(await manager.save(loaded)).toBe("saved");

            expect(await manager.load(created.uid)).toBeUndefined();
            expect(await manager.save(loaded)).toBe("missing");
        });

        it("Holds at most MAX_SESSIONS_PER_USER sessions per user, ending the oldest first, without touching other users.", async () => {
            const { manager } = build();
            const other = await manager.create("mailbox-2", "user-2");
            const sessions = [];
            for (let i = 0; i < MAX_SESSIONS_PER_USER; i++) {
                sessions.push(await manager.create("mailbox-1", "user-1"));
            }
            // A session the user already ended doesn't count against the limit.
            await manager.destroy(sessions[5].uid);
            const replacement = await manager.create("mailbox-1", "user-1");
            expect(await manager.load(sessions[0].uid)).toBeDefined();

            const overLimit = await manager.create("mailbox-1", "user-1");

            expect(await manager.load(sessions[0].uid)).toBeUndefined();
            expect(await manager.load(sessions[1].uid)).toBeDefined();
            expect(await manager.load(replacement.uid)).toBeDefined();
            expect(await manager.load(overLimit.uid)).toBeDefined();
            expect(await manager.load(other.uid)).toBeDefined();
        });
    });

    it("Redis store: never serves a stale per-process copy - a change saved through another replica is seen on the next load.", async () => {
        const client = new FakeRedisClient();
        const podA = redisManager(client);
        const podB = redisManager(client);
        const created = await podA.create("mailbox-1", "user-1");
        await podA.load(created.uid); // would have primed a local cache before

        const onB = (await podB.load(created.uid))!;
        onB.handles[3] = { type: "folder", entityUid: "folder:x" };
        expect(await podB.save(onB)).toBe("saved");

        expect((await podA.load(created.uid))!.handles[3]).toEqual({ type: "folder", entityUid: "folder:x" });
    });

    it("Redis store: saves through the compare-and-set script with the session TTL.", async () => {
        const client = new FakeRedisClient();
        const manager = redisManager(client);
        const created = await manager.create("mailbox-1", "user-1");
        const loaded = (await manager.load(created.uid))!;

        await manager.save(loaded);

        expect(client.eval).toHaveBeenCalledWith(expect.stringContaining("SETEX"), {
            keys: [`mapi.session.${created.uid}`],
            arguments: ["0", expect.any(String), String(SESSION_TTL_SECONDS)],
        });
        expect(client.ttls.get(`mapi.session.${created.uid}`)).toBe(SESSION_TTL_SECONDS);
    });

    it("RedisMapiSessionStore maps a missing key to undefined/false.", async () => {
        const store = new RedisMapiSessionStore(new FakeRedisClient());
        expect(await store.get("nope")).toBeUndefined();
        expect(await store.exists("nope")).toBe(false);
    });

    it("MemoryMapiSessionStore expires entries after their TTL, for reads, saves and later writes.", async () => {
        vi.useFakeTimers();
        const store = new MemoryMapiSessionStore();
        await store.put("a", JSON.stringify({ version: 0 }), 1);
        await store.put("b", JSON.stringify({ version: 0 }), 1);
        expect(await store.exists("a")).toBe(true);

        vi.advanceTimersByTime(1001);

        expect(await store.get("a")).toBeUndefined();
        expect(await store.compareAndSet("b", 0, "{}", 1)).toBe("missing");
        await store.put("c", "{}", 10); // sweeps the expired "b"
        expect(await store.exists("b")).toBe(false);
        expect(await store.exists("c")).toBe(true);
    });
});
