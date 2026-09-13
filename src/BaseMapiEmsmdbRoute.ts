///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ApiError, ObjectDecorators, type JWTUser } from "@rapidrest/core";
import {
    ApiErrorMessages,
    ApiErrors,
    HttpRequest,
    HttpResponse,
    ObjectFactory,
    RepoUtils,
    RouteDecorators,
} from "@rapidrest/service-core";
import { BufferReader, BufferWriter } from "./codec/BufferCursor.js";
import { decodeRopBuffer, encodeRopBuffer } from "./codec/RopBuffer.js";
import { MapiSessionContext, MapiSessionManager } from "./MapiSessionManager.js";
import { dispatchRops } from "./RopDispatcher.js";
import type { RopContext, RopHandler } from "./rop/RopHandler.js";
import { ScanPipeline } from "@rapidmx/restapi/scan";
import { Folder, Mailbox, RecoverableRepoUtils, resolveCallerMailboxUid, type BlobStore } from "@rapidmx/restapi";
const { Init, Inject, Logger } = ObjectDecorators;
const { Auth, Post, Request, Response, User: AuthUser } = RouteDecorators;

/** The well-known MAPI HRESULT `MAPI_E_LOGON_FAILED`, reused here to signal "no such session - reconnect"
 * via `X-ResponseCode`/`ErrorCode`. Not a claim of exact `[MS-OXCRPC]` return-value-table parity for this
 * specific condition - a real client only needs a non-zero code to know to re-`Connect`, not a precise one. */
const ERROR_SESSION_NOT_FOUND = 0x80040111;

function firstHeader(req: HttpRequest, name: string): string | undefined {
    const value: string | string[] | undefined = req.headers[name];
    return Array.isArray(value) ? value[0] : value;
}

/**
 * Abstract base for the single fixed EMSMDB endpoint (`POST /mapi/emsmdb` by `[MS-OXCMAPIHTTP]` convention,
 * though the concrete path is left to the consuming application to mount via `@Route(...)` - see
 * `BaseEasRoute.ts` for the identical undecorated-base-class pattern this follows). Like EAS, MAPI/HTTP
 * multiplexes several request types against one URL - here via the `X-RequestType` **header**
 * (`Connect`/`Execute`/`Disconnect`/`NotificationWait`) rather than a query parameter - so there is exactly
 * one `@Post()` method, not one per request type.
 *
 * **Auth**: `@Auth(["jwt"])` - unchanged from every other route in this app. Real, modern Exchange Server's
 * own MAPI virtual directory genuinely supports `OAuth` as a configured `IISAuthenticationMethods` value
 * alongside `NTLM`/`Negotiate`, so Bearer-token auth on this exact endpoint is real current Exchange
 * behavior, not a deviation this library invents - see the architecture plan's "Auth" section.
 *
 * **Connect/Disconnect are complete**: `Connect` establishes a real `MapiSessionContext` (via
 * `MapiSessionManager`) and sets the two spec-fixed session cookies (`MapiContext`/`MapiSequence`);
 * `Disconnect` releases the session. **`Execute` dispatches real ROPs** (starting with `RopLogon`/`RopRelease`
 * - build-order step 4) via `RopDispatcher`, using whichever `RopHandler`s `ropHandlerClasses` registers;
 * an unrecognized `RopId` simply stops processing (see `RopDispatcher`'s own doc comment for why), the same
 * "real, functional, no commands implemented yet" stance `BaseEasRoute`'s own skeleton step took for its
 * empty `commandHandlerClasses` before `ProvisionCommand` landed - here scoped per-ROP instead of per-request.
 *
 * **Always non-chunked**: every response here uses `Content-Length` (via `res.status(200).send(buffer)`),
 * never `Transfer-Encoding: chunked` - both are equally spec-valid per `[MS-OXCMAPIHTTP]`'s own "Common
 * Response Format", and skipping the `PROCESSING`/`PENDING`/`DONE` keep-alive meta-tag streaming matches
 * every other route in this library's own response idiom. See the architecture plan's "Protocol facts"
 * section for the full reasoning and its documented limitation.
 *
 * `mailboxClass`/`folderClass` are supplied by the Mongo/SQL concrete subclasses, following the exact
 * one-line-per-backend pattern used throughout this library's other routes/jobs. `folderRepo` is built once
 * here (not per-handler, unlike EAS's per-command repo pattern) and threaded through `RopContext` to every
 * `RopHandler` - see `RopHandler.ts`'s own doc comment for why.
 *
 * @author Jean-Philippe Steinmetz
 */
export abstract class BaseMapiEmsmdbRoute<M extends Mailbox> {
    protected abstract mailboxClass: any;
    protected abstract folderClass: any;
    protected abstract messageClass: any;
    protected abstract calendarEventClass: any;
    protected abstract contactClass: any;
    protected abstract taskClass: any;
    protected abstract labelClass: any;

    /** ROP handler classes to instantiate (one each) in `@Init`, keyed by their own `ropId`. Empty until a
     * concrete `RopHandler` lands - every ROP is then simply left unprocessed (see `RopDispatcher`'s own doc
     * comment), the correct, honest behavior for a transport skeleton with no ROPs implemented yet. */
    protected ropHandlerClasses: any[] = [];

    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private mailboxRepo?: RepoUtils<M>;
    private folderRepo?: RepoUtils<Folder>;
    private messageRepo?: RepoUtils<any>;
    private calendarEventRepo?: RepoUtils<any>;
    private contactRepo?: RepoUtils<any>;
    private taskRepo?: RepoUtils<any>;
    private labelRepo?: RepoUtils<any>;
    private sessionManager?: MapiSessionManager;
    private readonly ropHandlers = new Map<number, RopHandler>();

    @Inject("BlobStore")
    private blobStore?: BlobStore;

    @Inject(ScanPipeline)
    private scanPipeline?: ScanPipeline;

    @Inject("MailTransport")
    private mailTransport?: any;

    @Logger
    private logger: any;

    @Init
    public async init(): Promise<void> {
        this.mailboxRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.mailboxClass.name,
            args: [this.mailboxClass],
        });
        this.folderRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.folderClass.name,
            args: [this.folderClass],
        });
        this.messageRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.messageClass.name,
            args: [this.messageClass],
        });
        this.calendarEventRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.calendarEventClass.name,
            args: [this.calendarEventClass],
        });
        this.contactRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.contactClass.name,
            args: [this.contactClass],
        });
        this.taskRepo = await this._objectFactory!.newInstance(RecoverableRepoUtils, {
            name: this.taskClass.name,
            args: [this.taskClass],
        });
        // Label extends plain BaseEntity, not RecoverableBaseEntity (see restapi's own model doc comment) - no
        // soft-delete support to preserve, so a plain RepoUtils is correct here, matching mailboxRepo's own
        // choice for the identical reason.
        this.labelRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.labelClass.name,
            args: [this.labelClass],
        });
        this.sessionManager = await this._objectFactory!.newInstance(MapiSessionManager);
        for (const HandlerClass of this.ropHandlerClasses) {
            const handler: RopHandler = await this._objectFactory!.newInstance(HandlerClass);
            this.ropHandlers.set(handler.ropId, handler);
        }
    }

    @Auth(["jwt"])
    @Post()
    public async dispatch(
        @Request req: HttpRequest,
        @Response res: HttpResponse,
        @AuthUser user?: JWTUser,
    ): Promise<void> {
        if (
            !this.mailboxRepo ||
            !this.folderRepo ||
            !this.messageRepo ||
            !this.calendarEventRepo ||
            !this.contactRepo ||
            !this.taskRepo ||
            !this.labelRepo ||
            !this.sessionManager ||
            !this.blobStore ||
            !this.scanPipeline ||
            !this.mailTransport
        ) {
            throw new ApiError(ApiErrors.INTERNAL_ERROR, 500, ApiErrorMessages.INTERNAL_ERROR);
        }
        if (!user) {
            throw new ApiError(ApiErrors.AUTH_PERMISSION_FAILURE, 403, ApiErrorMessages.AUTH_PERMISSION_FAILURE);
        }

        const requestType: string | undefined = firstHeader(req, "x-requesttype");
        const clientInfo: string = firstHeader(req, "x-clientinfo") ?? "";
        const requestId: string = firstHeader(req, "x-requestid") ?? "";

        res.setHeader("Content-Type", "application/mapi-http")
            .setHeader("X-RequestType", requestType ?? "")
            .setHeader("X-RequestId", requestId)
            .setHeader("X-ResponseCode", "0")
            .setHeader("X-ClientInfo", clientInfo)
            .setHeader("X-ServerApplication", "RapidREST-Mail");

        switch (requestType) {
            case "Connect":
                await this.handleConnect(req, res, user);
                return;
            case "Execute":
                await this.handleExecute(req, res);
                return;
            case "Disconnect":
                await this.handleDisconnect(req, res);
                return;
            case "NotificationWait":
                // Real, spec-defined request type, deliberately deferred - no push-notification support in
                // this pass (see the architecture plan's ROP scope table). Distinct from an unrecognized
                // X-RequestType value (400 below), the same distinction BaseEasRoute draws between 501
                // ("recognized but deferred") and 400 ("malformed request").
                res.status(501).send();
                return;
            default:
                throw new ApiError(ApiErrors.INVALID_REQUEST, 400, ApiErrorMessages.INVALID_REQUEST);
        }
    }

    /**
     * `Connect` establishes a new Session Context - never trusts the request body's own `UserDn` field for
     * mailbox identity (the same "never a client-supplied mailbox, always resolved from the authenticated
     * JWT" principle `BaseEasRoute`/`BaseSearchRoute` already apply via `resolveCallerMailboxUid`) - so the
     * request body is never decoded at all here.
     */
    private async handleConnect(req: HttpRequest, res: HttpResponse, user: JWTUser): Promise<void> {
        const mailboxUid: string | undefined = await resolveCallerMailboxUid(this.mailboxRepo!, user);
        if (!mailboxUid) {
            throw new ApiError(ApiErrors.NOT_FOUND, 404, ApiErrorMessages.NOT_FOUND);
        }
        const mailbox: M | undefined = await this.mailboxRepo!.findOne(mailboxUid, { ignoreACL: true });

        const session: MapiSessionContext = await this.sessionManager!.create(mailboxUid, user.uid);
        res.appendHeader("Set-Cookie", `MapiContext=${session.uid}`);
        res.appendHeader("Set-Cookie", `MapiSequence=0`);

        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(60000); // PollsMax (ms) - pragmatic constant, not a tuned server policy value
        body.writeUInt32LE(3); // RetryCount - pragmatic constant
        body.writeUInt32LE(5000); // RetryDelay (ms) - pragmatic constant
        // DnPrefix: this library addresses mailboxes by UID/SMTP-address, never real X.500 DNs, so this is a
        // fixed, unused placeholder rather than a real DN prefix a client could build recipients from.
        body.writeNullTerminatedString8("/");
        body.writeNullTerminatedUtf16LE(mailbox?.displayName ?? "");
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }

    /**
     * `Execute` decodes the outer envelope (`Flags`/`RopBufferSize`/`RopBuffer`/...) and the inner ROP buffer
     * framing (`RopBuffer.ts`), dispatches every contained ROP via `RopDispatcher`, and re-encodes the
     * collected responses - preserving the incoming `handleTable` unchanged (this pragmatic subset never
     * allocates/frees table-wide handle slots at the framing level; individual `RopHandler`s manage their own
     * entries within `session.handles` instead).
     */
    private async handleExecute(req: HttpRequest, res: HttpResponse): Promise<void> {
        const sessionId: string | undefined = req.cookies["MapiContext"];
        const session: MapiSessionContext | undefined = sessionId
            ? await this.sessionManager!.load(sessionId)
            : undefined;
        if (!session) {
            res.setHeader("X-ResponseCode", String(ERROR_SESSION_NOT_FOUND));
            const body = new BufferWriter();
            body.writeUInt32LE(0); // StatusCode
            body.writeUInt32LE(ERROR_SESSION_NOT_FOUND); // ErrorCode
            body.writeUInt32LE(0); // Flags
            body.writeUInt32LE(0); // RopBufferSize
            body.writeUInt32LE(0); // AuxiliaryBufferSize
            res.status(200).send(body.toBuffer());
            return;
        }

        const requestReader = new BufferReader(req.rawBody ?? Buffer.alloc(0));
        requestReader.readUInt32LE(); // Flags - unused by this pragmatic subset (no client ROP-response hints honored)
        const ropBufferSize: number = requestReader.readUInt32LE();
        const ropBufferBytes: Buffer = requestReader.readBytes(ropBufferSize);
        // MaxRopOut/AuxiliaryBufferSize/AuxiliaryBuffer intentionally left unread - no output-size capping or
        // auxiliary-payload support in this pragmatic subset.

        const { ropsList, handleTable } = decodeRopBuffer(ropBufferBytes);
        const context: RopContext = {
            mailboxUid: session.mailboxUid,
            userUid: session.userUid,
            session,
            mailboxRepo: this.mailboxRepo!,
            folderRepo: this.folderRepo!,
            messageRepo: this.messageRepo!,
            calendarEventRepo: this.calendarEventRepo!,
            contactRepo: this.contactRepo!,
            taskRepo: this.taskRepo!,
            labelRepo: this.labelRepo!,
            folderClass: this.folderClass,
            messageClass: this.messageClass,
            calendarEventClass: this.calendarEventClass,
            contactClass: this.contactClass,
            taskClass: this.taskClass,
            labelClass: this.labelClass,
            blobStore: this.blobStore!,
            scanPipeline: this.scanPipeline!,
            mailTransport: this.mailTransport!,
        };
        const responseRopsList: Buffer = await dispatchRops(ropsList, this.ropHandlers, context);
        const responseRopBuffer: Buffer = encodeRopBuffer({ ropsList: responseRopsList, handleTable });

        await this.sessionManager!.save(session);

        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(0); // Flags
        body.writeUInt32LE(responseRopBuffer.length);
        body.writeBytes(responseRopBuffer);
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }

    private async handleDisconnect(req: HttpRequest, res: HttpResponse): Promise<void> {
        const sessionId: string | undefined = req.cookies["MapiContext"];
        if (sessionId) {
            await this.sessionManager!.destroy(sessionId);
        }
        const body = new BufferWriter();
        body.writeUInt32LE(0); // StatusCode
        body.writeUInt32LE(0); // ErrorCode
        body.writeUInt32LE(0); // AuxiliaryBufferSize
        res.status(200).send(body.toBuffer());
    }
}
