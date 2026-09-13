///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { RepoUtils } from "@rapidrest/service-core";
import type { BlobStore } from "@rapidmx/restapi";
import type { ScanPipeline } from "@rapidmx/restapi/scan";
import type { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import type { MapiSessionContext } from "../MapiSessionManager.js";

/**
 * Everything a `RopHandler` needs beyond the raw ROP bytes it decodes itself. `folderRepo`/`mailboxRepo` are
 * built once by `BaseMapiEmsmdbRoute` (not by each handler, unlike EAS's per-command repo pattern - MAPI's ROPs
 * are far more tightly interdependent around the same few entity types than EAS's per-command-collection-type
 * structure was, so centralizing the common repos here avoids every handler rebuilding its own). Untyped as
 * `RepoUtils<any>` rather than a generic `F extends Folder` parameter threaded through every handler file -
 * the same pragmatic `any` this codebase already uses throughout for `xxxClass`/`xxxRepo` DI fields, since
 * `RepoUtils<F>`'s method contravariance makes a generic version genuinely awkward for no real benefit here
 * (every handler only ever reads `Folder`-shaped rows, never constructs backend-specific entities itself).
 *
 * `folderClass`/`messageClass` (the concrete Mongo/SQL entity constructors) are threaded through here rather
 * than held as a field on each handler - unlike EAS's `EasCommandHandler`s, a single shared `RopHandler`
 * instance (e.g. `RopSubmitMessageHandler`) is reused for *both* the Mongo and SQL route (`ropHandlerClasses`
 * lists the exact same classes in `MapiEmsmdbRouteMongo`/`SQL`), so the concrete class can only ever be known
 * per-request, via this context - needed by `RopSubmitMessageHandler`'s `findOrCreateWellKnownFolder()`/
 * `messageRepo.create()` calls when persisting a submitted message's Sent Items copy. `scanPipeline`/
 * `mailTransport` back the same `scanAndRelay()` compose/send pipeline `BaseMessageRoute.send()`/EAS's
 * `ComposeMailCommand` already share via `MailSendUtils.ts`.
 *
 * `contactRepo`/`taskRepo`/`contactClass`/`taskClass`/`labelRepo`/`labelClass` are **optional**, unlike every
 * other repo/class pair here - `BaseMapiEmsmdbRoute` always populates them in real use (see its own `@Init`),
 * but making them required would force every one of this codebase's many existing `RopContext`-literal test
 * fixtures (most predating Contacts/Tasks folder support, and all predating Label support) to grow new fields
 * each for no behavioral reason. Every consumer (`PropertyResolvers.ts`, `RopGetContentsTableHandler.ts`,
 * `RopOpenMessageHandler.ts`) already degrades a `"contact:"`/`"task:"` target to empty/no-rows when the
 * corresponding repo is absent, the same "missing data, not a crash" stance `resolveContactInfo`/
 * `resolveTaskInfo` take for a vanished row; `labelRepo` absent degrades `PidNameKeywords` (Outlook Categories)
 * to an empty list the same way.
 */
export interface RopContext {
    mailboxUid: string;
    userUid: string;
    session: MapiSessionContext;
    mailboxRepo: RepoUtils<any>;
    folderRepo: RepoUtils<any>;
    messageRepo: RepoUtils<any>;
    calendarEventRepo: RepoUtils<any>;
    contactRepo?: RepoUtils<any>;
    taskRepo?: RepoUtils<any>;
    labelRepo?: RepoUtils<any>;
    folderClass: any;
    messageClass: any;
    calendarEventClass: any;
    contactClass?: any;
    taskClass?: any;
    labelClass?: any;
    blobStore: BlobStore;
    scanPipeline: ScanPipeline;
    mailTransport: any;
}

/**
 * One class per ROP, mirroring `EasCommandHandler`'s exact one-class-per-command convention. A handler owns
 * its own request/response wire format entirely - decoding whatever fields follow its own `RopId` byte from
 * `reader` (already consumed by the dispatcher) and appending its response bytes (if any - `RopRelease`
 * intentionally appends none) to the shared `writer`. See `RopBuffer.ts`'s doc comment for why a generic
 * dispatcher can't do this decoding itself.
 */
export interface RopHandler {
    /** The `RopId` byte this handler processes (e.g. `0xFE` for `RopLogon`). */
    readonly ropId: number;
    handle(reader: BufferReader, writer: BufferWriter, context: RopContext): Promise<void> | void;
}
