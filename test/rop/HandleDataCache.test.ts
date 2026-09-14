///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { HandleDataCache, handleDataKey } from "../../src/rop/HandleDataCache.js";

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

    it("Builds keys from session, handle index and generation.", () => {
        expect(handleDataKey("s1", 4, 9)).toBe("s1:4:9");
        expect(handleDataKey("s1", 4, undefined)).toBe("s1:4:0");
    });
});
