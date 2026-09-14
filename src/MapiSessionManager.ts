///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ObjectDecorators } from "@rapidrest/core";
import { DatabaseDecorators, SimpleEntity } from "@rapidrest/service-core";
import { handleDataCache, handleDataKey } from "./rop/HandleDataCache.js";
const { Init } = ObjectDecorators;
const { Redis } = DatabaseDecorators;

/** A pragmatic idle-session lifetime. The spec leaves session lifetime to server policy, not a fixed wire
 * value - a real client re-establishes (`Connect`) transparently whenever its session context has expired. */
export const SESSION_TTL_SECONDS = 15 * 60;

/** The longest a session may live, however active it stays. A client simply `Connect`s again afterwards. */
export const MAX_SESSION_LIFETIME_MS = 24 * 60 * 60 * 1000;

/** The most sessions one user may hold at once. `Connect` past this ends that user's oldest session. */
export const MAX_SESSIONS_PER_USER = 20;

/** Tags what a ROP-assigned integer handle (the `ServerObjectHandleTable` index space) refers to.
 * `entityUid` for a `"folder"` handle is one of `session.folderIds`' own value strings (`"virtual:<name>"` or
 * `"folder:<uid>"`), not a bare UID - the same format throughout avoids a second parallel encoding.
 *
 * A `"table"` handle's `columns`/`cursor` hold `RopSetColumns`/`RopQueryRows` state. A hierarchy table carries its
 * resolved child-folder targets in `rows`. A contents table carries no rows at all: `contentsKind` says which repo
 * its folder's items live in and `RopQueryRows` reads just the requested window from the database (see
 * `ContentsTable.ts`), so a large folder never lands in session state.
 *
 * A `"stream"` handle's `entityUid` is the `"message:<uid>"` target its content was opened from,
 * `propertyId`/`propertyType` the `PropertyTag` `RopOpenStream` opened (this pragmatic subset only ever supports
 * `PidTagBody`/`PtypString`, see `MessageBodyStream.ts`), and `streamPosition` how many bytes `RopReadStream` has
 * already returned. A read stream's decoded body lives in `HandleDataCache`, not here. A write stream (opened
 * `ReadWrite`/`Create` against a `RopCreateMessage` draft's `PidTagBody`) records the draft it belongs to in
 * `writeTargetHandleIndex`/`writeTargetGeneration`, and its accumulated bytes in `writeBufferBase64`/`writeSize`.
 *
 * A `"message"` handle from `RopCreateMessage` (a draft not yet `RopSaveChangesMessage`d) has `entityUid: ""`
 * and instead carries `draftFolderUid` (the folder it will belong to) and `draftProperties` (the small,
 * well-known set of properties this pragmatic subset's `RopSetProperties` tracks - Subject/DisplayTo/
 * DisplayCc/DisplayBcc/an inline `PidTagBody`, each coerced to a plain string, keyed by decimal `PropertyId` -
 * a string key because a JSON-object key is always a string regardless of how it's written). Buffers are stored
 * as base64 and dates as ISO strings since sessions round-trip through `JSON.stringify`/`JSON.parse`.
 *
 * A `"fastTransfer"` handle (`RopFastTransferSourceCopyTo`/`CopyProperties`'s output handle) keeps only what is
 * needed to rebuild its stream (`transferSourceType`, `transferColumns`/`transferExcludeIds`) plus the paging
 * cursor `transferPosition`; the built stream itself lives in `HandleDataCache`.
 *
 * `generation` is unique per assignment within a session (see `assignHandle`), so data keyed by handle index can
 * tell a handle apart from a later one that reuses the same index. */
export interface MapiObjectHandle {
    type: "logon" | "folder" | "message" | "table" | "stream" | "fastTransfer";
    entityUid: string;
    generation?: number;
    rows?: string[];
    contentsKind?: "message" | "calendarEvent" | "contact" | "task";
    columns?: { propertyId: number; propertyType: number }[];
    cursor?: number;
    propertyId?: number;
    propertyType?: number;
    streamPosition?: number;
    draftFolderUid?: string;
    draftProperties?: Record<string, string>;
    writeTargetHandleIndex?: number;
    writeTargetGeneration?: number;
    writeBufferBase64?: string;
    writeSize?: number;
    transferSourceType?: "folder" | "message";
    transferColumns?: { propertyId: number; propertyType: number }[];
    transferExcludeIds?: number[];
    transferPosition?: number;
}

/**
 * The MAPI/HTTP `Session Context` (`[MS-OXCMAPIHTTP]` §3.1.1.1): everything a `Connect`-established session
 * needs across subsequent `Execute` requests. Never persisted to a real database - purely an ephemeral,
 * TTL-bound store entry (see `MapiSessionManager` below), so this deliberately extends `SimpleEntity` (just a
 * `uid`) rather than this library's own `BaseEntity`.
 *
 * `createdAt` is a plain ISO-8601 string, not a `Date`, since sessions round-trip through JSON. `version` is
 * bumped on every successful save and compared on the next one - see `MapiSessionManager.save()`.
 */
export class MapiSessionContext extends SimpleEntity {
    public mailboxUid: string;
    public userUid: string;
    public version = 0;
    public handles: Record<number, MapiObjectHandle> = {};
    public nextHandleIndex = 1;
    public nextHandleGeneration = 1;
    public createdAt: string = new Date().toISOString();
    /** This session's FID assignments for the 13 `RopLogon` special folders, keyed by FID (decimal string),
     * valued `"virtual:<name>"` or `"folder:<uid>"` - see `RopLogonHandler`'s own doc comment. Populated by
     * `RopLogon`, read back by a later `RopOpenFolder`. */
    public folderIds: Record<string, string> = {};

    /** The reverse of `folderIds` (target -> FID) plus a monotonic counter, maintained alongside it by
     * `FolderTarget.assignOrGetFid`/`RopLogonHandler` so "does this target already have a FID" is an O(1)
     * lookup instead of a linear scan of `folderIds` repeated once per table row. */
    public folderTargetIds: Record<string, number> = {};
    public nextFolderId = 1;

    /** This session's MID assignments, keyed by MID (decimal string), valued `"message:<uid>"` - the message
     * analog of `folderIds` above. A MID only ever comes into existence lazily, the first time a `RopQueryRows`
     * row exposes a message's `PidTagMid` column (see `MessageTarget.assignOrGetMid`), read back by a later
     * `RopOpenMessage`. */
    public messageIds: Record<string, string> = {};

    /** The reverse of `messageIds` (target -> MID) plus a monotonic counter. */
    public messageTargetIds: Record<string, number> = {};
    public nextMessageId = 1;

    /** This session's `RopGetPropertyIdsFromNames` mapping table (`[MS-OXCPRPT]` §2.2.12), keyed by a JSON
     * string encoding of the `{guid, kind, lid|name}` `PropertyName` the numeric ID was assigned to - see
     * `NamedPropertyRegistry.ts`. */
    public namedProperties: Record<string, number> = {};

    /** The reverse of `namedProperties` (numeric ID -> `PropertyName` JSON key) plus a monotonic counter. */
    public namedPropertyIds: Record<number, string> = {};
    public nextNamedPropertyId = 0x8000;

    public constructor(other: Partial<SimpleEntity> & { mailboxUid: string; userUid: string }) {
        super(other);
        this.mailboxUid = other.mailboxUid;
        this.userUid = other.userUid;
    }
}

/**
 * Stores `handle` at `index` with a fresh generation. Whatever previously held that index is released first
 * (see `releaseHandle`), so a stream or cached data belonging to the old handle never attaches to the new one.
 */
export function assignHandle(session: MapiSessionContext, index: number, handle: MapiObjectHandle): MapiObjectHandle {
    releaseHandle(session, index);
    handle.generation = session.nextHandleGeneration++;
    session.handles[index] = handle;
    return handle;
}

/**
 * Releases the handle at `index`: drops its cached data, and releases every write stream opened against it
 * (a stream is meaningless once its draft message is gone, and would otherwise attach to whatever reuses the
 * index next).
 */
export function releaseHandle(session: MapiSessionContext, index: number): void {
    const handle = session.handles[index];
    if (!handle) {
        return;
    }
    delete session.handles[index];
    handleDataCache.delete(handleDataKey(session.uid, index, handle.generation));
    for (const [key, candidate] of Object.entries(session.handles)) {
        if (candidate.writeTargetHandleIndex === index && candidate.writeTargetGeneration === handle.generation) {
            releaseHandle(session, Number(key));
        }
    }
}

/** The outcome of `MapiSessionManager.save()`. `"conflict"` means another request saved the session first;
 * `"missing"` that it expired or was ended. */
export type SessionSaveResult = "saved" | "conflict" | "missing";

/** The storage a `MapiSessionManager` runs on: Redis when a `cache` datastore is configured, otherwise memory. */
export interface MapiSessionStore {
    get(key: string): Promise<string | undefined>;
    exists(key: string): Promise<boolean>;
    put(key: string, value: string, ttlSeconds: number): Promise<void>;
    compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<SessionSaveResult>;
    delete(key: string): Promise<void>;
}

/** Atomically replaces the session only when its stored `version` still equals the expected one. Returns 1 when
 * saved, 0 on a version mismatch and -1 when the key no longer exists. */
const COMPARE_AND_SET_SCRIPT = `
local current = redis.call('GET', KEYS[1])
if not current then return -1 end
local ok, decoded = pcall(cjson.decode, current)
if not ok or tonumber(decoded['version']) ~= tonumber(ARGV[1]) then return 0 end
redis.call('SETEX', KEYS[1], ARGV[3], ARGV[2])
return 1
`;

/** A `MapiSessionStore` on a node-redis client. Every read goes to Redis, never a per-process copy, so each
 * replica sees the same session. */
export class RedisMapiSessionStore implements MapiSessionStore {
    public constructor(private readonly client: any) {}

    public async get(key: string): Promise<string | undefined> {
        return (await this.client.get(key)) ?? undefined;
    }

    public async exists(key: string): Promise<boolean> {
        return (await this.client.exists(key)) > 0;
    }

    public async put(key: string, value: string, ttlSeconds: number): Promise<void> {
        await this.client.setEx(key, ttlSeconds, value);
    }

    public async compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<SessionSaveResult> {
        const result = Number(
            await this.client.eval(COMPARE_AND_SET_SCRIPT, { keys: [key], arguments: [String(expectedVersion), value, String(ttlSeconds)] }),
        );
        return result === 1 ? "saved" : result === 0 ? "conflict" : "missing";
    }

    public async delete(key: string): Promise<void> {
        await this.client.del(key);
    }
}

/** A single-process `MapiSessionStore`. Values are stored as JSON strings so every load hands out an independent
 * copy, exactly like the Redis store. */
export class MemoryMapiSessionStore implements MapiSessionStore {
    private readonly entries = new Map<string, { value: string; expiresAt: number }>();

    public async get(key: string): Promise<string | undefined> {
        const entry = this.entries.get(key);
        if (entry && entry.expiresAt <= Date.now()) {
            this.entries.delete(key);
            return undefined;
        }
        return entry?.value;
    }

    public async exists(key: string): Promise<boolean> {
        return (await this.get(key)) !== undefined;
    }

    public async put(key: string, value: string, ttlSeconds: number): Promise<void> {
        this.sweep();
        this.entries.set(key, { value, expiresAt: Date.now() + ttlSeconds * 1000 });
    }

    public async compareAndSet(key: string, expectedVersion: number, value: string, ttlSeconds: number): Promise<SessionSaveResult> {
        // No await between the read and the write, so no other request can interleave.
        const entry = this.entries.get(key);
        if (!entry || entry.expiresAt <= Date.now()) {
            return "missing";
        }
        if (JSON.parse(entry.value).version !== expectedVersion) {
            return "conflict";
        }
        this.entries.set(key, { value, expiresAt: Date.now() + ttlSeconds * 1000 });
        return "saved";
    }

    public async delete(key: string): Promise<void> {
        this.entries.delete(key);
    }

    private sweep(): void {
        const now = Date.now();
        for (const [key, entry] of this.entries) {
            if (entry.expiresAt <= now) {
                this.entries.delete(key);
            }
        }
    }
}

function sessionKey(sessionId: string): string {
    return `mapi.session.${encodeURIComponent(sessionId)}`;
}

function userIndexKey(userUid: string): string {
    return `mapi.session.user.${encodeURIComponent(userUid)}`;
}

/**
 * Stores/loads `MapiSessionContext`s, keyed by the opaque session id that also becomes the `MapiContext` cookie
 * value.
 *
 * Every load reads the shared store (Redis when a `cache` datastore is configured). An earlier version used
 * `RedisCache`, whose per-process copy served a session without checking Redis, so with several replicas one pod
 * could run ROPs against handles another pod had already changed. Saves are compare-and-set on `version`: when
 * two requests on the same session overlap, the second save reports `"conflict"` instead of overwriting the
 * first request's handles.
 */
export class MapiSessionManager {
    @Redis("cache", false)
    private redisClient?: any;

    private store?: MapiSessionStore;

    @Init
    public init(): void {
        this.store = this.redisClient ? new RedisMapiSessionStore(this.redisClient) : new MemoryMapiSessionStore();
    }

    /** Creates a session for `userUid`, ending that user's oldest sessions if they already hold
     * `MAX_SESSIONS_PER_USER`. */
    public async create(mailboxUid: string, userUid: string): Promise<MapiSessionContext> {
        const context = new MapiSessionContext({ mailboxUid, userUid });
        await this.store!.put(sessionKey(context.uid), JSON.stringify(context), SESSION_TTL_SECONDS);

        const indexKey = userIndexKey(userUid);
        const indexed: string[] = JSON.parse((await this.store!.get(indexKey)) ?? "[]");
        const live: string[] = [];
        for (const id of indexed) {
            if (await this.store!.exists(sessionKey(id))) {
                live.push(id);
            }
        }
        while (live.length >= MAX_SESSIONS_PER_USER) {
            await this.destroy(live.shift()!);
        }
        live.push(context.uid);
        await this.store!.put(indexKey, JSON.stringify(live), MAX_SESSION_LIFETIME_MS / 1000);
        return context;
    }

    /** Loads a session straight from the store. A session past `MAX_SESSION_LIFETIME_MS` is ended and treated as
     * missing. */
    public async load(sessionId: string): Promise<MapiSessionContext | undefined> {
        const json = await this.store!.get(sessionKey(sessionId));
        if (!json) {
            return undefined;
        }
        const session: MapiSessionContext = JSON.parse(json);
        if (Date.now() - Date.parse(session.createdAt) > MAX_SESSION_LIFETIME_MS) {
            await this.destroy(sessionId);
            return undefined;
        }
        return session;
    }

    /** Saves `context` only if nobody else saved it since it was loaded, bumping its `version` on success. */
    public async save(context: MapiSessionContext): Promise<SessionSaveResult> {
        const expected = context.version;
        context.version = expected + 1;
        const result = await this.store!.compareAndSet(sessionKey(context.uid), expected, JSON.stringify(context), SESSION_TTL_SECONDS);
        if (result !== "saved") {
            context.version = expected;
        }
        return result;
    }

    public async destroy(sessionId: string): Promise<void> {
        await this.store!.delete(sessionKey(sessionId));
    }
}
