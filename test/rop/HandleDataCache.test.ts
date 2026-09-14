///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { defaultHandleDataStore, HandleDataCache, handleDataKey, MemoryHandleDataStore, RedisHandleDataStore } from "../../src/rop/HandleDataCache.js";
import { handleDataStoreOf } from "../../src/rop/RopHandler.js";
import { FakeRedisClient } from "../fakeRedis.js";

describe("HandleDataCache Tests", () => {
    afterEach(() => {
        vi.useRealTimers();
    });

    it("Stores, returns and deletes entries, tracking the bytes held.", () => {
        const cache = new HandleDataCache(100, 1000);
        cache.set("a", Buffer.alloc(10));
        expect(cache.get("a")?.length).toBe(10);
        expect(cache.size).toBe(10);

        cache.set("a", Buffer.alloc(4)); // replacing an entry doesn't double-count it
        expect(cache.size).toBe(4);

        cache.delete("a");
        cache.delete("missing");
        expect(cache.get("a")).toBeUndefined();
        expect(cache.size).toBe(0);
    });

    it("Evicts least recently used entries once the byte budget is exceeded.", () => {
        const cache = new HandleDataCache(30, 1000);
        cache.set("a", Buffer.alloc(10));
        cache.set("b", Buffer.alloc(10));
        cache.set("c", Buffer.alloc(10));
        cache.get("a"); // "b" is now the least recently used

        cache.set("d", Buffer.alloc(10));

        expect(cache.get("b")).toBeUndefined();
        expect(cache.get("a")).toBeDefined();
        expect(cache.get("c")).toBeDefined();
        expect(cache.get("d")).toBeDefined();
        expect(cache.size).toBe(30);
    });

    it("Doesn't cache a value larger than the whole budget.", () => {
        const cache = new HandleDataCache(8, 1000);
        cache.set("big", Buffer.alloc(9));
        expect(cache.get("big")).toBeUndefined();
        expect(cache.size).toBe(0);
    });

    it("Expires an entry that hasn't been read within the TTL.", () => {
        vi.useFakeTimers();
        const cache = new HandleDataCache(100, 1000);
        cache.set("a", Buffer.alloc(5));
        vi.advanceTimersByTime(999);
        expect(cache.get("a")).toBeDefined(); // reading refreshes the TTL
        vi.advanceTimersByTime(999);
        expect(cache.get("a")).toBeDefined();
        vi.advanceTimersByTime(1000);
        expect(cache.get("a")).toBeUndefined();
        expect(cache.size).toBe(0);
    });

    it("Builds keys from session, handle index and generation nonce.", () => {
        expect(handleDataKey("s1", 4, "nonce")).toBe("s1:4:nonce");
        expect(handleDataKey("s1", 4, undefined)).toBe("s1:4:");
    });
});

describe.each([
    ["memory", () => new MemoryHandleDataStore(new HandleDataCache())],
    ["redis", () => new RedisHandleDataStore(new FakeRedisClient(), new HandleDataCache())],
])("HandleDataStore on %s", (_name, build) => {
    it("Stores, returns and deletes values.", async () => {
        const store = build();
        expect(await store.get("k")).toBeUndefined();
        await store.set("k", Buffer.from("value"), 60);
        expect(await store.get("k")).toEqual(Buffer.from("value"));
        await store.delete("k");
        expect(await store.get("k")).toBeUndefined();
    });

    it("Reassembles a write stream from offset-keyed chunks, a rewritten offset replacing its chunk.", async () => {
        const store = build();
        await store.putChunk("w", 0, Buffer.from("abc"), 60);
        await store.putChunk("w", 3, Buffer.from("XX"), 60); // a write whose request lost its session save...
        await store.putChunk("w", 3, Buffer.from("de"), 60); // ...retried at the same offset
        await store.putChunk("w", 5, Buffer.from("f"), 60);

        expect(await store.readChunks("w", 6)).toEqual(Buffer.from("abcdef"));
        expect(await store.readChunks("w", 5)).toEqual(Buffer.from("abcde"));
        expect(await store.readChunks("w", 0)).toEqual(Buffer.alloc(0));
        // A size that doesn't land on a chunk boundary, or a missing chunk, can't be reassembled exactly.
        expect(await store.readChunks("w", 4)).toBeUndefined();
        expect(await store.readChunks("w", 7)).toBeUndefined();
        expect(await store.readChunks("missing", 3)).toBeUndefined();
    });
});

describe("RedisHandleDataStore", () => {
    it("Shares values between replicas through Redis, with the TTL, and fills the local cache on read.", async () => {
        const client = new FakeRedisClient();
        const podA = new RedisHandleDataStore(client, new HandleDataCache());
        const localB = new HandleDataCache();
        const podB = new RedisHandleDataStore(client, localB);

        await podA.set("s:5:n", Buffer.from("stream"), 600);

        expect(client.ttls.get("mapi.handle.s:5:n")).toBe(600);
        expect(await podB.get("s:5:n")).toEqual(Buffer.from("stream"));
        expect(localB.get("s:5:n")).toEqual(Buffer.from("stream"));
        client.values.clear();
        expect(await podB.get("s:5:n")).toEqual(Buffer.from("stream")); // served locally now
    });

    it("Treats a client returning no hash as an empty stream.", async () => {
        const client = new FakeRedisClient();
        client.hGetAll = vi.fn().mockResolvedValue(null) as any;
        expect(await new RedisHandleDataStore(client).readChunks("w", 1)).toBeUndefined();
    });
});

describe("handleDataStoreOf", () => {
    it("Uses the context's store, falling back to the process-wide memory store.", () => {
        const store = new MemoryHandleDataStore();
        expect(handleDataStoreOf({ handleData: store })).toBe(store);
        expect(handleDataStoreOf({})).toBe(defaultHandleDataStore);
    });
});
