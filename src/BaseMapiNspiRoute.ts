///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ApiError, ObjectDecorators, type JWTUser } from "@rapidrest/core";
import { ApiErrorMessages, ApiErrors, HttpRequest, HttpResponse, ObjectFactory, RepoUtils, RouteDecorators } from "@rapidrest/service-core";
import { handleNspiBind, handleNspiUnbind } from "./nspi/NspiBindHandler.js";
import { handleNspiGetMatches } from "./nspi/NspiGetMatchesHandler.js";
import { Mailbox, resolveCallerMailboxUid } from "@rapidmx/restapi";
const { Init } = ObjectDecorators;
const { Auth, Post, Request, Response, User: AuthUser } = RouteDecorators;

function firstHeader(req: HttpRequest, name: string): string | undefined {
    const value: string | string[] | undefined = req.headers[name];
    return Array.isArray(value) ? value[0] : value;
}

/**
 * Abstract base for the single fixed NSPI endpoint (`POST /mapi/nspi` by `[MS-OXCMAPIHTTP]` convention, though
 * the concrete path is left to the consuming application to mount via `@Route(...)`, the identical
 * undecorated-base-class pattern `BaseMapiEmsmdbRoute`/`BaseEasRoute` already establish). Dispatches on the
 * `X-RequestType` header, the same multiplexing mechanism EMSMDB uses - but unlike EMSMDB, there is no ROP
 * buffer framing here: each `X-RequestType` value (`Bind`/`Unbind`/`GetMatches`) is its own flat, self-
 * contained request/response body (`[MS-OXCMAPIHTTP]` §2.2.5), so this route dispatches directly to a small
 * handler function per operation instead of decoding a multiplexed stream.
 *
 * **Pragmatic subset scope**: only `Bind`/`Unbind`/`GetMatches` are implemented - real NSPI's other dozen-plus
 * request types (`QueryRows`, `ResolveNames`, `GetProps`, `ModProps`, ...) are genuine full address-book/
 * directory-replication operations this deployment's own minimal GAL-lookup use case doesn't need; an
 * unrecognized `X-RequestType` gets a `501`, the same "recognized-but-deferred vs malformed" distinction
 * `BaseMapiEmsmdbRoute`'s own `NotificationWait` case already draws. See `NspiBindHandler.ts`'s own doc comment
 * for why no real NSPI session state is created or tracked between calls.
 *
 * **Auth**: `@Auth(["jwt"])`, unchanged from every other route in this app - the same reasoning
 * `BaseMapiEmsmdbRoute`'s own doc comment already gives for why Bearer-token auth on an address-book endpoint
 * is real, current Exchange behavior, not a deviation this library invents.
 *
 * `mailboxClass`/`contactClass` are supplied by the Mongo/SQL concrete subclasses, and `likePattern` picks the
 * two-backend `like()`-wrapping convention `SearchCommand.ts` (EAS) already established for its own identical
 * GAL-search query gap.
 *
 * @author Jean-Philippe Steinmetz
 */
export abstract class BaseMapiNspiRoute<M extends Mailbox> {
    protected abstract mailboxClass: any;
    protected abstract contactClass: any;

    /** Wraps an already-regex-escaped substring for this backend's `like()` operator - see
     * `NspiGetMatchesHandler.ts`'s own `findMatchingContacts` and `SearchCommand.ts`'s identical hook for the
     * full Mongo-vs-SQL reasoning. */
    protected abstract likePattern(escaped: string): string;

    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private mailboxRepo?: RepoUtils<M>;
    private contactRepo?: RepoUtils<any>;

    @Init
    public async init(): Promise<void> {
        this.mailboxRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.mailboxClass.name,
            args: [this.mailboxClass],
        });
        this.contactRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.contactClass.name,
            args: [this.contactClass],
        });
    }

    @Auth(["jwt"])
    @Post()
    public async dispatch(@Request req: HttpRequest, @Response res: HttpResponse, @AuthUser user?: JWTUser): Promise<void> {
        if (!this.mailboxRepo || !this.contactRepo) {
            throw new ApiError(ApiErrors.INTERNAL_ERROR, 500, ApiErrorMessages.INTERNAL_ERROR);
        }
        if (!user) {
            throw new ApiError(ApiErrors.AUTH_PERMISSION_FAILURE, 403, ApiErrorMessages.AUTH_PERMISSION_FAILURE);
        }

        const requestType: string | undefined = firstHeader(req, "x-requesttype");
        res.setHeader("Content-Type", "application/mapi-http")
            .setHeader("X-RequestType", requestType ?? "")
            .setHeader("X-ResponseCode", "0")
            .setHeader("X-ServerApplication", "RapidREST-Mail");

        switch (requestType) {
            case "Bind":
                await handleNspiBind(res, user, this.mailboxRepo);
                return;
            case "Unbind":
                handleNspiUnbind(res);
                return;
            case "GetMatches": {
                const mailboxUid = await resolveCallerMailboxUid(this.mailboxRepo, user);
                if (!mailboxUid) {
                    throw new ApiError(ApiErrors.NOT_FOUND, 404, ApiErrorMessages.NOT_FOUND);
                }
                await handleNspiGetMatches(req, res, mailboxUid, this.contactRepo, (escaped) => this.likePattern(escaped));
                return;
            }
            default:
                res.status(501).send();
                return;
        }
    }
}
