///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Isolated unit tests for BaseMapiNspiRoute, reserved for the defensive guard branches and the unrecognized-
// X-RequestType dispatch fallback a real wired server exercises only rarely - the same rationale
// test/routes/BaseMapiEmsmdbRoute.test.ts already uses for its own identical guards. Bind/Unbind/GetMatches'
// real behavior is exercised via real HTTP+DB requests in test/routes/mongo/MapiNspiRoute.test.ts (and its
// sql/ counterpart).
import config from "../config.js";
import { ObjectFactory } from "@rapidrest/service-core";
import { Logger } from "@rapidrest/core";
import { BaseMapiNspiRoute } from "../../src/BaseMapiNspiRoute.js";

class TestMapiNspiRoute extends BaseMapiNspiRoute<any> {
    protected mailboxClass: any = { name: "TestMailbox" };
    protected contactClass: any = { name: "TestContact" };

    protected likePattern(term: string): string {
        return term;
    }
}

function makeReq(requestType?: string): any {
    return { headers: requestType ? { "x-requesttype": requestType } : {}, cookies: {}, rawBody: undefined };
}

function makeRes(): any {
    return {
        status: vi.fn().mockReturnThis(),
        setHeader: vi.fn().mockReturnThis(),
        appendHeader: vi.fn().mockReturnThis(),
        send: vi.fn().mockReturnThis(),
    };
}

describe("BaseMapiNspiRoute Tests (guard clauses + dispatch fallback only)", () => {
    const objectFactory: ObjectFactory = new ObjectFactory(config, Logger());

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("dispatch() throws INTERNAL_ERROR when mailboxRepo/contactRepo are not set.", async () => {
        // `initialize: false` skips `@Init`, leaving both genuinely `undefined` - exactly what this guard
        // clause exists to catch.
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });

        await expect(route.dispatch(makeReq("Bind"), makeRes(), { uid: "user-1" } as any)).rejects.toThrow(/internal error/i);
    });

    it("dispatch() throws AUTH_PERMISSION_FAILURE when no authenticated user is present.", async () => {
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });
        (route as any).mailboxRepo = {};
        (route as any).contactRepo = {};

        await expect(route.dispatch(makeReq("Bind"), makeRes(), undefined)).rejects.toThrow(/permission/i);
    });

    it("dispatch() returns 501 for an unrecognized X-RequestType.", async () => {
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });
        (route as any).mailboxRepo = {};
        (route as any).contactRepo = {};
        const res = makeRes();

        await route.dispatch(makeReq("SomeUnsupportedOperation"), res, { uid: "user-1" } as any);

        expect(res.status).toHaveBeenCalledWith(501);
    });

    it("dispatch() returns 501 and echoes an empty X-RequestType when the header is missing entirely.", async () => {
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });
        (route as any).mailboxRepo = {};
        (route as any).contactRepo = {};
        const res = makeRes();

        await route.dispatch(makeReq(), res, { uid: "user-1" } as any);

        expect(res.setHeader).toHaveBeenCalledWith("X-RequestType", "");
        expect(res.status).toHaveBeenCalledWith(501);
    });

    it("dispatch() takes the first value when X-RequestType is sent as multiple header values.", async () => {
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });
        (route as any).mailboxRepo = { find: vi.fn().mockResolvedValue([]) };
        (route as any).contactRepo = {};
        const res = makeRes();
        const req = makeReq();
        req.headers["x-requesttype"] = ["Bind", "Unbind"];

        await route.dispatch(req, res, { uid: "user-1" } as any);

        expect(res.setHeader).toHaveBeenCalledWith("X-RequestType", "Bind");
    });

    it("dispatch() throws NOT_FOUND for GetMatches when the caller owns no mailbox.", async () => {
        const route = objectFactory.newInstance<TestMapiNspiRoute>(TestMapiNspiRoute, { initialize: false });
        (route as any).mailboxRepo = { find: vi.fn().mockResolvedValue([]) };
        (route as any).contactRepo = {};

        await expect(route.dispatch(makeReq("GetMatches"), makeRes(), { uid: "user-1" } as any)).rejects.toThrow(/no resource could be found/i);
    });
});
