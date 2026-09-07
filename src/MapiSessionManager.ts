///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ObjectDecorators } from "@rapidrest/core";
import { ObjectFactory, RedisCache, SimpleEntity } from "@rapidrest/service-core";
const { Init } = ObjectDecorators;

/** A pragmatic idle-session lifetime. The spec leaves session lifetime to server policy, not a fixed wire
 * value - a real client re-establishes (`Connect`) transparently whenever its session context has expired. */
const SESSION_TTL_SECONDS = 15 * 60;

/** Tags what a ROP-assigned integer handle (the `ServerObjectHandleTable` index space) refers to.
 * `entityUid` for a `"folder"` handle is one of `session.folderIds`' own value strings (`"virtual:<name>"` or
 * `"folder:<uid>"`), not a bare UID - the same format throughout avoids a second parallel encoding. A
 * `"table"` handle's `rows`/`columns`/`cursor` hold `RopGetHierarchyTable`/`RopSetColumns`/`RopQueryRows`
 * state for that specific table instance: `rows` is the resolved, order-fixed list of entity targets (same
 * `"virtual:<name>"`/`"folder:<uid>"` format) this table enumerates, `columns` the `RopSetColumns`-configured
 * property list, `cursor` how many rows `RopQueryRows` has already returned. A `"stream"` handle's `entityUid`
 * is the `"message:<uid>"` target its content was opened from, `propertyId`/`propertyType` the `PropertyTag`
 * `RopOpenStream` opened (this pragmatic subset only ever supports `PidTagBody`/`PtypString`, see
 * `MessageBodyStream.ts`), and `streamPosition` how many bytes `RopReadStream` has already returned (a
 * read-mode stream) or `writeTargetHandleIndex`/`writeBufferBase64` the accumulated write state (a write-mode
 * stream opened `ReadWrite`/`Create` against a `RopCreateMessage` draft's `PidTagBody` - see
 * `RopOpenStreamHandler`'s own doc comment).
 *
 * A `"message"` handle from `RopCreateMessage` (a draft not yet `RopSaveChangesMessage`d) has `entityUid: ""`
 * and instead carries `draftFolderUid` (the folder it will belong to) and `draftProperties` (the small,
 * well-known set of properties this pragmatic subset's `RopSetProperties` tracks - Subject/DisplayTo/
 * DisplayCc/DisplayBcc/an inline `PidTagBody`, each coerced to a plain string, keyed by decimal `PropertyId` -
 * a string key because a JSON-object key is always a string regardless of how it's written). `writeBufferBase64`
 * stores accumulated `RopWriteStream` bytes as base64 rather than a raw `Buffer` for the same reason `Date`
 * fields elsewhere in this class are stored as ISO strings: `RedisCache`'s Redis-backed path round-trips
 * everything through `JSON.stringify`/`JSON.parse`, which cannot represent a `Buffer` (or a `bigint`, which is
 * why no property value is ever stored in its native decoded MAPI type here) losslessly.
 *
 * A `"fastTransfer"` handle (`RopFastTransferSourceCopyTo`/`CopyProperties`'s output handle) carries the
 * entire pre-built FastTransfer stream eagerly (`transferBufferBase64`, same base64-for-`RedisCache` reasoning
 * as `writeBufferBase64`) - this pragmatic subset builds the whole stream up front rather than truly
 * incrementally, since a real client only ever pages it out via `RopFastTransferSourceGetBuffer`'s own
 * `transferPosition` cursor into that already-complete buffer, see `FastTransferStream.ts`'s own doc comment. */
export interface MapiObjectHandle {
    type: "logon" | "folder" | "message" | "table" | "stream" | "fastTransfer";
    entityUid: string;
    rows?: string[];
    columns?: { propertyId: number; propertyType: number }[];
    cursor?: number;
    propertyId?: number;
    propertyType?: number;
    streamPosition?: number;
    draftFolderUid?: string;
    draftProperties?: Record<string, string>;
    writeTargetHandleIndex?: number;
    writeBufferBase64?: string;
    transferBufferBase64?: string;
    transferPosition?: number;
}

/**
 * The MAPI/HTTP `Session Context` (`[MS-OXCMAPIHTTP]` §3.1.1.1): everything a `Connect`-established session
 * needs across subsequent `Execute` requests. Never persisted to a real database - purely an ephemeral,
 * TTL-bound cache entry (see `MapiSessionManager` below), so this deliberately extends `SimpleEntity` (just a
 * `uid`) rather than this library's own `BaseEntity`, which would add ACL/soft-delete/optimistic-locking
 * machinery this object has no use for.
 *
 * `createdAt` is a plain ISO-8601 string, not a `Date` - `RedisCache`'s Redis-backed path round-trips values
 * through `JSON.stringify`/`JSON.parse`, which silently turns a `Date` into a string on the way out without
 * reviving it back to a `Date` on the way in; storing it as a string from the start keeps the shape identical
 * whether a session happens to be served from the in-memory or the Redis-backed path.
 */
export class MapiSessionContext extends SimpleEntity {
    public mailboxUid: string;
    public userUid: string;
    public handles: Record<number, MapiObjectHandle> = {};
    public nextHandleIndex = 1;
    public createdAt: string = new Date().toISOString();
    /** This session's FID assignments for the 13 `RopLogon` special folders, keyed by FID (decimal string),
     * valued `"virtual:<name>"` or `"folder:<uid>"` - see `RopLogonHandler`'s own doc comment. Populated by
     * `RopLogon`, read back by a later `RopOpenFolder`. */
    public folderIds: Record<string, string> = {};

    /** This session's MID assignments, keyed by MID (decimal string), valued `"message:<uid>"` - the message
     * analog of `folderIds` above. Unlike `folderIds` (pre-populated by `RopLogon`), a MID only ever comes into
     * existence lazily, the first time a `RopQueryRows` row exposes a message's `PidTagMid` column (see
     * `MessageTarget.assignOrGetMid`), read back by a later `RopOpenMessage`. */
    public messageIds: Record<string, string> = {};

    /** This session's `RopGetPropertyIdsFromNames` mapping table (`[MS-OXCPRPT]` §2.2.12), keyed by a JSON
     * string encoding of the `{guid, kind, lid|name}` `PropertyName` the numeric ID was assigned to - see
     * `NamedPropertyRegistry.ts`'s own doc comment for why a JSON string (rather than a delimiter-joined one)
     * is the safest key shape here. A real client resolves every named property (almost every Calendar-specific
     * one - `PidLidAppointmentStartWhole`, `PidLidBusyStatus`, `PidLidAppointmentRecur`, ...) through this table
     * once per session before ever setting/reading it via `RopSetProperties`/`RopGetPropertiesSpecific`. */
    public namedProperties: Record<string, number> = {};

    /** `mailboxUid`/`userUid` are always known at construction time (the only call site is
     * `MapiSessionManager.create()`, which resolves both up front) - required here rather than optional with
     * a same-value fallback, which would just be dead code no real caller ever takes the other branch of. */
    public constructor(other: Partial<SimpleEntity> & { mailboxUid: string; userUid: string }) {
        super(other);
        this.mailboxUid = other.mailboxUid;
        this.userUid = other.userUid;
    }
}

/**
 * Stores/loads `MapiSessionContext`s, keyed by the opaque session id that also becomes the `MapiContext`
 * cookie value. Reuses `@rapidrest/service-core`'s existing generic `RedisCache<T>` (Redis-backed when a
 * `cache` datastore is configured, transparent in-memory fallback otherwise) rather than either a bespoke
 * Redis client (unlike `PingCommand`'s own throwaway pub/sub client - fine for ephemeral fan-out, not a good
 * fit for stateful session data) or the framework's generic cookie/session middleware (built for a whole-app
 * `req.session` object, not this protocol's specific handle-table shape). `RedisCache` needs no real
 * persisted entity - only a class with a `.name` for its Redis key prefix - so a purely ephemeral,
 * never-persisted class like `MapiSessionContext` is a perfectly valid `type` for it.
 *
 * @author Jean-Philippe Steinmetz
 */
export class MapiSessionManager {
    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private cache?: RedisCache<MapiSessionContext>;

    @Init
    public async init(): Promise<void> {
        this.cache = await this._objectFactory!.newInstance(RedisCache, {
            name: MapiSessionContext.name,
            args: [MapiSessionContext],
        });
    }

    public async create(mailboxUid: string, userUid: string): Promise<MapiSessionContext> {
        const context = new MapiSessionContext({ mailboxUid, userUid });
        await this.cache!.save(context.uid, context, SESSION_TTL_SECONDS);
        return context;
    }

    public async load(sessionId: string): Promise<MapiSessionContext | undefined> {
        return this.cache!.load(sessionId);
    }

    public async save(context: MapiSessionContext): Promise<void> {
        await this.cache!.save(context.uid, context, SESSION_TTL_SECONDS);
    }

    public async destroy(sessionId: string): Promise<void> {
        await this.cache!.delete(sessionId);
    }
}
