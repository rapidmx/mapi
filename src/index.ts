///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/**
 * MAPI-over-HTTP protocol compatibility (Phase 3 of this library's roadmap): a hand-built ROP (Remote
 * Operation) codec and buffer-framing layer (`[MS-OXCROPS]`), JWT-secured EMSMDB transport (`BaseMapiEmsmdbRoute`,
 * reusing the framework's existing `@Auth(["jwt"])` strategy - no MAPI-specific auth code of its own), and a
 * minimal NSPI address-book endpoint (`BaseMapiNspiRoute`) for GAL "search as you type" lookups.
 *
 * This module exports only the backend-agnostic surface: the wire-format codecs, the abstract route base
 * classes, `MapiSessionManager`, the ROP dispatcher, and every ROP/NSPI handler class (each takes a generic
 * `RopContext`/repo argument, not a concrete entity class, so none of them are backend-specific). The concrete
 * Mongo/SQL classes a deployment actually instantiates (`MapiEmsmdbRouteMongo`/`SQL`,
 * `MapiNspiRouteMongo`/`SQL`) are exported from this package's `./mongo`/`./sql` subpaths instead, alongside
 * every other entity/route/job this library defines - see `src/eas/index.ts`'s identical doc comment for the
 * same convention.
 *
 * A deployment mounts both MAPI endpoints with a trivial one-line subclass each, the same pattern
 * `src/eas/index.ts`'s doc comment describes for `BaseEasRoute`:
 * ```ts
 * import { MapiEmsmdbRouteMongo, MapiNspiRouteMongo } from "@rapidrest/mail/mongo";
 * import { RouteDecorators } from "@rapidrest/service-core";
 * const { Route } = RouteDecorators;
 *
 * @Route("/mapi/emsmdb")
 * export class MyMapiEmsmdbRoute extends MapiEmsmdbRouteMongo {}
 *
 * @Route("/mapi/nspi")
 * export class MyMapiNspiRoute extends MapiNspiRouteMongo {}
 * ```
 * Real Outlook desktop locates both URLs via `@rapidrest/mail/autodiscover`'s Outlook/EXCH response
 * (`BaseAutodiscoverRoute`'s `mapiUrl` property) - see that module's own doc comment.
 *
 * **Documented gaps in this pragmatic subset** (see the architecture plan's "MAPI over HTTP (Phase 3)" section
 * for the full reasoning behind each):
 * - Incremental sync (`RopFastTransferSourceCopyTo`/`CopyProperties`/`GetBuffer`) always builds a full,
 *   non-differential dump, not real IDSET-based ICS (`[MS-OXCFXICS]`) - a real client still works correctly
 *   against this, just less efficiently than byte-perfect ICS.
 * - Calendar/meetings: no counter-proposals, meeting-forwarding, delegate scheduling, or resource-booking
 *   auto-accept; no DST-aware timezones (fixed-offset approximation only); no recurrence exceptions/modified
 *   instances; no Hijri/`MonthEnd` recurrence patterns.
 * - No `RopModifyRecipients` - attendees/recipients are always added via the same `PidTagDisplayTo`/`Cc`
 *   pragmatic path used for ordinary mail.
 * - No public-folder support, no delegate/shared-mailbox access, no rules/permissions/search-folder ROPs, no
 *   client-certificate enrollment.
 * - Calendar named properties (location/recurrence/timezone/etc.) are excluded from FastTransfer streams -
 *   only the columns in `FastTransferStream.ts`'s `DEFAULT_*_COLUMNS` tables are included.
 * - NSPI is limited to `Bind`/`Unbind`/`GetMatches` only - no row/name lookups beyond a single content-
 *   restriction search, no directory replication.
 */
export * from "./codec/BufferCursor.js";
export * from "./codec/MapiGuid.js";
export * from "./codec/TypedString.js";
export * from "./codec/PropertyValue.js";
export * from "./codec/RopBuffer.js";
export * from "./codec/AppointmentRecurrence.js";
export * from "./codec/MapiTimeZone.js";
export * from "./codec/GlobalObjectId.js";
export * from "./RopDispatcher.js";
export * from "./MapiSessionManager.js";
export * from "./BaseMapiEmsmdbRoute.js";
export * from "./BaseMapiNspiRoute.js";
export * from "./rop/RopHandler.js";
export * from "./rop/FolderTarget.js";
export * from "./rop/MessageTarget.js";
export * from "./rop/CalendarEventTarget.js";
export * from "./rop/MessageBodyStream.js";
export * from "./rop/PropertyResolvers.js";
export * from "./rop/NamedPropertyRegistry.js";
export * from "./rop/CalendarNamedProperties.js";
export * from "./rop/FastTransferStream.js";
export * from "./rop/MeetingMessageClassHandler.js";
export * from "./rop/RopLogonHandler.js";
export * from "./rop/RopReleaseHandler.js";
export * from "./rop/RopOpenFolderHandler.js";
export * from "./rop/RopGetHierarchyTableHandler.js";
export * from "./rop/RopSetColumnsHandler.js";
export * from "./rop/RopQueryRowsHandler.js";
export * from "./rop/RopGetContentsTableHandler.js";
export * from "./rop/RopOpenMessageHandler.js";
export * from "./rop/RopGetPropertiesSpecificHandler.js";
export * from "./rop/RopOpenStreamHandler.js";
export * from "./rop/RopReadStreamHandler.js";
export * from "./rop/RopCreateMessageHandler.js";
export * from "./rop/RopSetPropertiesHandler.js";
export * from "./rop/RopWriteStreamHandler.js";
export * from "./rop/RopSaveChangesMessageHandler.js";
export * from "./rop/RopSubmitMessageHandler.js";
export * from "./rop/RopGetPropertyIdsFromNamesHandler.js";
export * from "./rop/RopDeleteMessagesHandler.js";
export * from "./rop/RopDeleteFolderHandler.js";
export * from "./rop/RopFastTransferSourceCopyToHandler.js";
export * from "./rop/RopFastTransferSourceCopyPropertiesHandler.js";
export * from "./rop/RopFastTransferSourceGetBufferHandler.js";
export * from "./nspi/NspiCodec.js";
export * from "./nspi/NspiBindHandler.js";
export * from "./nspi/NspiGetMatchesHandler.js";
