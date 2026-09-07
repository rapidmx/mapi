///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import * as http from "http";

export interface MapiRawResponse {
    status: number;
    headers: http.IncomingHttpHeaders;
    body: Buffer;
}

function resolvePort(app: any): number {
    if (typeof app === "number") {
        return app;
    }
    if (app && typeof app.port === "number") {
        return app.port;
    }
    if (app && typeof app.listenPort === "number") {
        return app.listenPort;
    }
    throw new Error("mapiRequest(): cannot determine port from argument");
}

/**
 * A minimal raw-socket HTTP client for MAPI/HTTP integration tests, deliberately bypassing the shared
 * `@rapidrest/service-core/test` `request()` helper. That helper reads every response via axios with
 * `responseType: "text"`, decoding the body through Node's default UTF-8 text decoding before handing it
 * back - lossless only when every response byte is below `0x80`. That's true for EAS's WBXML payloads (its
 * own multi-byte integer encoding is 7-bit-safe by design) but NOT true for MAPI/HTTP's raw little-endian
 * 32-bit integers and binary structures, which routinely contain bytes >= `0x80` (e.g. `60000` = `0xEA60`;
 * byte `0xEA` is an invalid leading UTF-8 byte with no valid continuation byte following it, corrupting
 * everything decoded after that point). Confirmed by a real failing test during this phase's own development,
 * not a theoretical concern - every MAPI/HTTP test in this phase uses this helper instead.
 */
export function mapiRequest(
    app: any,
    path: string,
    headers: Record<string, string>,
    body: Buffer,
): Promise<MapiRawResponse> {
    const port = resolvePort(app);
    return new Promise((resolve, reject) => {
        const req = http.request(
            {
                hostname: "localhost",
                port,
                path,
                method: "POST",
                headers: { ...headers, "Content-Length": body.length },
            },
            (res) => {
                const chunks: Buffer[] = [];
                res.on("data", (chunk: Buffer) => chunks.push(chunk));
                res.on("end", () =>
                    resolve({ status: res.statusCode ?? 0, headers: res.headers, body: Buffer.concat(chunks) }),
                );
            },
        );
        req.on("error", reject);
        req.end(body);
    });
}

/** Extracts just the `name=value` pairs from a `Set-Cookie` response header array/string, joined for reuse as
 * a single request `Cookie` header - real cookie attributes (`Path`/`HttpOnly`/...) aren't relevant here. */
export function cookieHeaderFrom(setCookie: string | string[] | undefined): string {
    if (!setCookie) {
        return "";
    }
    const values = Array.isArray(setCookie) ? setCookie : [setCookie];
    return values.map((c) => c.split(";")[0]).join("; ");
}
