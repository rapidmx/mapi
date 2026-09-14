///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { MailTransport, OutboundMessage, TransportResult } from "@rapidmx/restapi";

/**
 * Sends `message` and throws unless the transport accepted it for at least one recipient and rejected none.
 *
 * The same rule as restapi's own `sendOrThrow()` (`transport/TransportResultUtils.ts`), which isn't exported from
 * the `@rapidmx/restapi` package root, so it can't be imported here; keep the two in sync. Every bundled transport
 * reports a failed relay through `TransportResult.rejected` rather than throwing, so a bare `send()` would treat a
 * rejected message as sent.
 */
export async function sendOrThrow(transport: MailTransport, message: OutboundMessage): Promise<TransportResult> {
    const result: TransportResult | undefined = await transport.send(message);
    if (!result || (result.accepted ?? []).length === 0 || (result.rejected ?? []).length > 0) {
        throw new Error(`The mail transport did not accept the message (rejected: ${(result?.rejected ?? []).length}).`);
    }
    return result;
}
