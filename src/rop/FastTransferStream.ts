///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { Folder, FolderType } from "@rapidmx/restapi";
import { BufferWriter } from "../codec/BufferCursor.js";
import { PropertyType, writeTaggedPropertyValue } from "../codec/PropertyValue.js";
import { assignHandle, type MapiObjectHandle } from "../MapiSessionManager.js";
import { resolveFolderCalendarEvents } from "./CalendarEventTarget.js";
import { handleDataKey } from "./HandleDataCache.js";
import { resolveFolderMessages } from "./MessageTarget.js";
import { resolvePropertyValues } from "./PropertyResolvers.js";
import { handleDataStoreOf, type RopContext } from "./RopHandler.js";

/**
 * Builds the `FastTransfer` binary stream (`[MS-OXCFXICS]` §2.2.4) `RopFastTransferSourceCopyTo`/
 * `CopyProperties` hand off to `RopFastTransferSourceGetBuffer` for paging out, per this pragmatic subset's own
 * "ICS" scope (see the architecture plan's "Incremental sync" row): a **full, non-differential** dump of the
 * target's current live state, built eagerly and stored whole on the output handle - not the real, much larger
 * `contentsSync`/`hierarchySync` grammar (`IncrSyncChg`/`IncrSyncDel`/`IncrSyncStateBegin` markers), which
 * requires a persisted per-device ICS state object (an `IDSET` baseline) this pragmatic subset doesn't
 * implement. A real client that asks for an incremental sync here simply gets a full re-list every time and
 * reconciles by diffing against its own local cache - spec-legitimate client behavior, just more state
 * re-fetched per sync than a byte-perfect ICS server would send, exactly the tradeoff the plan accepts.
 *
 * Only the plain `PidTag*` properties this pragmatic subset already resolves via `PropertyResolvers.ts` are
 * ever included - a Calendar item's own named properties (start/end/location/recurrence/...) are **not**
 * included in the stream (only `PidTagSubject`), a deliberate scope cut: a client wanting those already has to
 * `RopOpenMessage`+`RopGetPropertiesSpecific` the item directly (the same "just re-fetches more state" tradeoff
 * applied one level deeper). No recipients/attachments/subfolders are emitted either - `messageChildren` is
 * always empty and `Level` (whether to recurse into subfolders) is not honored, matching this pragmatic
 * subset's existing "no `RopModifyRecipients`" and folder-hierarchy-is-flat-per-call precedents elsewhere.
 *
 * @author Jean-Philippe Steinmetz
 */

/** FastTransfer markers (`[MS-OXCFXICS]` §2.2.4.1.4, confirmed this session) - only the four this pragmatic
 * subset's flat, non-differential dump ever needs. */
const MARKER_START_TOP_FLD = 0x40090003;
const MARKER_END_FOLDER = 0x400b0003;
const MARKER_START_MESSAGE = 0x400c0003;
const MARKER_END_MESSAGE = 0x400d0003;

interface PropertyColumn {
    propertyId: number;
    propertyType: PropertyType;
}

/** The well-known columns this pragmatic subset includes by default when the caller (`RopFastTransferSourceCopyTo`)
 * didn't supply an explicit include list - the same small "what a real client needs" sets `PropertyResolvers.ts`
 * already defines its own defaults around. */
const DEFAULT_FOLDER_COLUMNS: PropertyColumn[] = [
    { propertyId: 0x3001, propertyType: PropertyType.PtypString }, // PidTagDisplayName
    { propertyId: 0x3602, propertyType: PropertyType.PtypInteger32 }, // PidTagContentCount
    { propertyId: 0x3603, propertyType: PropertyType.PtypInteger32 }, // PidTagContentUnreadCount
];
const DEFAULT_MESSAGE_COLUMNS: PropertyColumn[] = [
    { propertyId: 0x0037, propertyType: PropertyType.PtypString }, // PidTagSubject
    { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 }, // PidTagMessageFlags
    { propertyId: 0x0e1b, propertyType: PropertyType.PtypBoolean }, // PidTagHasAttachments
    { propertyId: 0x0e06, propertyType: PropertyType.PtypTime }, // PidTagMessageDeliveryTime
];
/** A calendar item's named properties are deliberately excluded from the stream (see class doc comment) -
 * `PidTagSubject` is the only column emitted for one. */
const DEFAULT_CALENDAR_COLUMNS: PropertyColumn[] = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }];

async function writePropList(
    writer: BufferWriter,
    target: string,
    columns: PropertyColumn[],
    context: Pick<RopContext, "mailboxUid" | "session" | "folderRepo" | "messageRepo" | "calendarEventRepo" | "mailboxRepo" | "budget">,
): Promise<void> {
    const values = await resolvePropertyValues(target, columns, context);
    columns.forEach((column, index) =>
        writeTaggedPropertyValue(writer, { propertyId: column.propertyId, propertyType: column.propertyType, value: values[index] }),
    );
}

function filterExcluded(columns: PropertyColumn[], excludePropertyIds: ReadonlySet<number>): PropertyColumn[] {
    return columns.filter((column) => !excludePropertyIds.has(column.propertyId));
}

/** The largest FastTransfer stream built for one handle; a larger source fails with `MAPI_E_TOO_BIG`. */
export const MAX_FAST_TRANSFER_BYTES = 32 * 1024 * 1024;

/** How long a built stream is kept for paging out. A client pages a stream out right after opening it; one that
 * leaves it longer than this has to open the transfer again. */
export const FAST_TRANSFER_TTL_SECONDS = 10 * 60;

/** Thrown while building a stream that has grown past its byte limit. */
export class FastTransferTooBigError extends Error {
    public constructor() {
        super("FastTransfer: the stream is larger than MAX_FAST_TRANSFER_BYTES.");
        this.name = "FastTransferTooBigError";
    }
}

/**
 * Builds the complete FastTransfer stream for `handle` (a `"folder"` or `"message"` Server object). `columns`,
 * when given (a `RopFastTransferSourceCopyProperties` explicit include list), replaces every default column set
 * uniformly at both folder- and message-level; `excludePropertyIds` (a `RopFastTransferSourceCopyTo` exclude
 * list) is only ever applied to this pragmatic subset's own default columns, per the same reasoning.
 *
 * The size is checked after every item, so a huge folder stops being built as soon as it passes `maxBytes`
 * (`FastTransferTooBigError`) instead of being fully built first. Every item resolved and the finished stream's bytes
 * count against the request's `ExecuteBudget` when there is one.
 */
export async function buildFastTransferStream(
    handle: MapiObjectHandle,
    context: RopContext,
    options: { columns?: PropertyColumn[]; excludePropertyIds?: ReadonlySet<number>; maxBytes?: number } = {},
): Promise<Buffer> {
    const writer = new BufferWriter();
    const exclude = options.excludePropertyIds ?? new Set<number>();
    const maxBytes = options.maxBytes ?? MAX_FAST_TRANSFER_BYTES;
    const checkSize = (): void => {
        if (writer.length > maxBytes) {
            throw new FastTransferTooBigError();
        }
    };

    if (handle.type === "folder") {
        const folderColumns = options.columns ?? filterExcluded(DEFAULT_FOLDER_COLUMNS, exclude);
        writer.writeUInt32LE(MARKER_START_TOP_FLD);
        await writePropList(writer, handle.entityUid, folderColumns, context);

        if (handle.entityUid.startsWith("folder:")) {
            const folderUid = handle.entityUid.slice("folder:".length);
            const folder: Folder | undefined = await context.folderRepo.findOne(folderUid, { ignoreACL: true });
            const isCalendar = folder?.type === FolderType.CALENDAR;
            const rows = isCalendar
                ? await resolveFolderCalendarEvents(folderUid, context.calendarEventRepo)
                : await resolveFolderMessages(folderUid, context.messageRepo);
            const messageColumns = options.columns ?? filterExcluded(isCalendar ? DEFAULT_CALENDAR_COLUMNS : DEFAULT_MESSAGE_COLUMNS, exclude);

            for (const row of rows) {
                writer.writeUInt32LE(MARKER_START_MESSAGE);
                await writePropList(writer, row, messageColumns, context);
                writer.writeUInt32LE(MARKER_END_MESSAGE);
                checkSize();
            }
        }

        writer.writeUInt32LE(MARKER_END_FOLDER);
    } else {
        const isCalendar = handle.entityUid.startsWith("calendarEvent:");
        const columns = options.columns ?? filterExcluded(isCalendar ? DEFAULT_CALENDAR_COLUMNS : DEFAULT_MESSAGE_COLUMNS, exclude);
        await writePropList(writer, handle.entityUid, columns, context);
    }
    checkSize();
    context.budget?.chargeBytes(writer.length);

    return writer.toBuffer();
}

/** Builds a `"fastTransfer"` handle's stream from what the handle records about its source, or `undefined` when it
 * is larger than `MAX_FAST_TRANSFER_BYTES`. */
async function rebuildFastTransferStream(transfer: MapiObjectHandle, context: RopContext): Promise<Buffer | undefined> {
    try {
        return await buildFastTransferStream({ type: transfer.transferSourceType!, entityUid: transfer.entityUid }, context, {
            columns: transfer.transferColumns,
            excludePropertyIds: new Set(transfer.transferExcludeIds ?? []),
        });
    } catch (err) {
        if (err instanceof FastTransferTooBigError) {
            return undefined;
        }
        throw err;
    }
}

/** Why `loadFastTransferBuffer` has no stream: `"tooBig"` when a rebuild passed `MAX_FAST_TRANSFER_BYTES`, `"lost"`
 * when the stored stream is gone part-way through paging. */
export type FastTransferLoadFailure = "tooBig" | "lost";

/**
 * The built stream for the `"fastTransfer"` handle at `handleIndex`, from the shared `HandleDataStore` (Redis when
 * configured, so any replica can page it out).
 *
 * When the stream is gone (expired, evicted, or built by a replica without shared storage), it is rebuilt only if
 * nothing has been paged out yet. Part-way through, a rebuild could differ from what the client already has (items
 * added or changed since), and continuing at the old offset would hand it bytes from a different stream, so that
 * case reports `"lost"` and the client restarts the transfer. A rebuild is held to `MAX_FAST_TRANSFER_BYTES` like
 * the original build.
 */
export async function loadFastTransferBuffer(
    context: RopContext,
    handleIndex: number,
    transfer: MapiObjectHandle,
): Promise<Buffer | FastTransferLoadFailure> {
    const store = handleDataStoreOf(context);
    const key = handleDataKey(context.session.uid, handleIndex, transfer.generation);
    const stored = await store.get(key);
    if (stored) {
        return stored;
    }
    if ((transfer.transferPosition ?? 0) > 0) {
        return "lost";
    }
    const buffer = await rebuildFastTransferStream(transfer, context);
    if (!buffer) {
        return "tooBig";
    }
    await store.set(key, buffer, FAST_TRANSFER_TTL_SECONDS);
    return buffer;
}

/**
 * Shared by `RopFastTransferSourceCopyTo`/`CopyProperties`: builds the stream for `source`, and on success stores a
 * `"fastTransfer"` handle at `outputHandleIndex` with the stream in the `HandleDataStore`. Returns `false`, storing
 * nothing, when the stream exceeds `MAX_FAST_TRANSFER_BYTES`.
 */
export async function openFastTransferHandle(
    context: RopContext,
    outputHandleIndex: number,
    source: MapiObjectHandle,
    options: { columns?: PropertyColumn[]; excludePropertyIds?: number[] },
): Promise<boolean> {
    const transfer: MapiObjectHandle = {
        type: "fastTransfer",
        entityUid: source.entityUid,
        transferSourceType: source.type === "folder" ? "folder" : "message",
        transferColumns: options.columns,
        transferExcludeIds: options.excludePropertyIds,
        transferPosition: 0,
    };
    const buffer = await rebuildFastTransferStream(transfer, context);
    if (!buffer) {
        return false;
    }
    assignHandle(context.session, outputHandleIndex, transfer);
    await handleDataStoreOf(context).set(handleDataKey(context.session.uid, outputHandleIndex, transfer.generation), buffer, FAST_TRANSFER_TTL_SECONDS);
    return true;
}
