///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { simpleParser } from "mailparser";
import type { RepoUtils } from "@rapidrest/service-core";
import { Message, type BlobStore } from "@rapidmx/restapi";
import { BufferWriter } from "../codec/BufferCursor.js";

/** `PidTagBody` (`[MS-OXPROPS]`): property ID `0x1000`, `PtypString`. The only streamable property this
 * pragmatic subset supports via `RopOpenStream`/`RopReadStream` - `PidTagHtml`/`PidTagRtfCompressed` are a
 * documented gap (see `RopOpenStreamHandler`'s own doc comment). */
export const PID_TAG_BODY = 0x1000;

/**
 * Resolves a `"message:<uid>"` target's plain-text body as the exact byte sequence `RopReadStream` must serve
 * for a streamed `PtypString` property - confirmed via `[MS-OXCPRPT]`'s `RopOpenStream` page ("a string of
 * Unicode characters in UTF-16LE format encoding with terminating null character"), the identical encoding
 * `writePropertyValue()` already uses for an inline `PtypString` value, so this reuses
 * `BufferWriter.writeNullTerminatedUtf16LE` rather than reimplementing it.
 *
 * Always derived from the message's raw MIME source (`bodyBlobKey`, via `mailparser` - the same fallback path
 * `ItemOperationsCommand.fetchMessage()` already uses for EAS's own plain-text body delivery) rather than
 * `sanitizedHtmlBlobKey` - `PidTagBody` is specifically the plain-text body per `[MS-OXPROPS]`, unlike EAS's
 * `Body` element, which can carry either format tagged by its own `Type` field.
 *
 * Recomputed on every call rather than cached anywhere (including on the `"stream"` handle itself) -
 * `MapiSessionContext` is serialized through `RedisCache`'s JSON round trip for multi-instance deployments, so
 * a raw `Buffer` stored there wouldn't survive it (the same class of gap this codebase already documented for
 * `Date` fields, see `MapiSessionContext`'s own doc comment). A documented, pragmatic trade-off: a large body
 * read across several `RopReadStream` calls re-parses the MIME source each time rather than once - correct,
 * not byte-perfect-efficient.
 */
export async function resolveMessageBodyBytes(target: string, messageRepo: RepoUtils<any>, blobStore: BlobStore): Promise<Buffer> {
    const uid = target.slice("message:".length);
    const message: Message | undefined = await messageRepo.findOne(uid, { ignoreACL: true });
    if (!message) {
        return Buffer.alloc(0);
    }
    const raw = await blobStore.get(message.bodyBlobKey);
    const parsed = await simpleParser(raw);
    const writer = new BufferWriter();
    writer.writeNullTerminatedUtf16LE(parsed.text ?? "");
    return writer.toBuffer();
}
