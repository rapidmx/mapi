///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// These tests prove BaseMapiEmsmdbRoute's transport plumbing (JWT auth, mailbox resolution, session-cookie
// establishment/teardown, the Execute envelope + RopBuffer framing round trip) AND the RopLogon/RopRelease/
// RopOpenFolder/RopGetHierarchyTable/RopGetContentsTable/RopSetColumns/RopQueryRows/RopOpenMessage/
// RopGetPropertiesSpecific/RopOpenStream/RopReadStream/RopGetPropertyIdsFromNames ROPs' real behavior, all over
// a real HTTP round trip - see the architecture plan's "Phase 3" section, build-order steps 3-9a.
//
// Uses `mapiTestClient.ts`'s raw-socket `mapiRequest()` instead of the shared `@rapidrest/service-core/test`
// `request()` helper - see that file's doc comment for why (the shared helper's axios `responseType: "text"`
// reading corrupts binary response bytes >= 0x80, which MAPI/HTTP's raw integers routinely contain).
import config from "../../config.js";
import { MongoConnection, MongoRepository, Server, ObjectFactory, ConnectionManager, ACLAction } from "@rapidrest/service-core";
import { JWTUtils, Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { ContactMongo, FolderMongo, MailboxMongo, MessageMongo, TaskMongo } from "@rapidmx/restapi/mongo";
import { FolderType, MessageImportance, RecipientType } from "@rapidmx/restapi";
import { MongoMemoryServer } from "mongodb-memory-server";
import { InMemoryBlobStore, RecordingMailTransport, registerTestDoubles } from "../../testDoubles.js";
import { cookieHeaderFrom, mapiRequest } from "../../mapiTestClient.js";
import { BufferReader, BufferWriter } from "../../../src/codec/BufferCursor.js";
import { decodeGuid, encodeGuid } from "../../../src/codec/MapiGuid.js";
import { PropertyType, readPropertyValue, writePropertyTag, writeTaggedPropertyValue, type TaggedPropertyValue } from "../../../src/codec/PropertyValue.js";
import { decodeRopBuffer, encodeRopBuffer } from "../../../src/codec/RopBuffer.js";

const mongod: MongoMemoryServer = new MongoMemoryServer({
    instance: {
        port: 9999,
        dbName: "rrst-test",
    },
});

describe("Route:MapiEmsmdbRouteMongo Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-mongo", logger, objectFactory });
    const baseUrl = "/mongo/mapi/emsmdb";
    let mailboxRepo: MongoRepository<MailboxMongo>;
    let folderRepo: MongoRepository<FolderMongo>;
    let messageRepo: MongoRepository<MessageMongo>;
    let contactRepo: MongoRepository<ContactMongo>;
    let taskRepo: MongoRepository<TaskMongo>;
    let aclRepo: MongoRepository<any>;

    const owner: any = { uid: uuid.v4(), roles: [], elevated: Date.now() };
    const ownerToken = JWTUtils.createTokenSync(config.get("auth"), owner);

    const createMailbox = async function (ownerUid: string, data?: Partial<MailboxMongo>): Promise<MailboxMongo> {
        const obj: MailboxMongo = new MailboxMongo({
            ownerUserUid: ownerUid,
            primarySmtpAddress: `${uuid.v4()}@example.com`,
            aliasAddresses: [],
            displayName: "Test Mailbox",
            timezone: "UTC",
            quotaBytes: 1_000_000_000,
            usedBytes: 0,
            ...data,
        });
        const result: MailboxMongo = await mailboxRepo.save(obj);
        await aclRepo.save({
            uid: result.uid,
            dateCreated: new Date(),
            dateModified: new Date(),
            version: 0,
            records: [{ userOrRoleId: ownerUid, actions: [ACLAction.FULL] }],
            parentUid: "Mailbox",
        });
        return result;
    };

    const createFolder = async function (mailboxUid: string, type: FolderType, name: string, parentFolderUid?: string): Promise<FolderMongo> {
        return await folderRepo.save(
            new FolderMongo({ mailboxUid, name, type, parentFolderUid, unreadCount: 0, totalCount: 0, syncKeyVersion: 0 }),
        );
    };

    const createMessage = async function (mailboxUid: string, folderUid: string, data?: Partial<MessageMongo>): Promise<MessageMongo> {
        return await messageRepo.save(
            new MessageMongo({
                mailboxUid,
                folderUid,
                messageId: `${uuid.v4()}@example.com`,
                subject: "Test Subject",
                from: { address: "sender@example.com", displayName: "Sender", type: RecipientType.TO },
                recipients: [{ address: "owner@example.com", type: RecipientType.TO }],
                sentDate: new Date(),
                receivedDate: new Date("2026-03-15T09:00:00.000Z"),
                bodyBlobKey: `bodies/${uuid.v4()}`,
                bodyPreview: "Hello world",
                flags: { read: true, flagged: false, answered: false, forwarded: false },
                importance: MessageImportance.NORMAL,
                references: [],
                hasAttachments: false,
                ...data,
            }),
        );
    };

    const createContact = async function (mailboxUid: string, folderUid: string, data?: Partial<ContactMongo>): Promise<ContactMongo> {
        return await contactRepo.save(
            new ContactMongo({ mailboxUid, folderUid, displayName: "Test Contact", emails: [], phones: [], addresses: [], ...data }),
        );
    };

    const createTask = async function (mailboxUid: string, folderUid: string, data?: Partial<TaskMongo>): Promise<TaskMongo> {
        return await taskRepo.save(new TaskMongo({ mailboxUid, folderUid, title: "Test Task", ...data }));
    };

    const blobStore = function (): InMemoryBlobStore {
        return objectFactory.getInstance<InMemoryBlobStore>("BlobStore")!;
    };

    const mailTransport = function (): RecordingMailTransport {
        return objectFactory.getInstance<RecordingMailTransport>("MailTransport")!;
    };

    const connect = async function (headers: Record<string, string> = {}) {
        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Connect",
                "Content-Type": "application/mapi-http",
                ...headers,
            },
            Buffer.alloc(0),
        );
    };

    const execute = async function (cookie: string, ropBuffer: Buffer) {
        const body = new BufferWriter();
        body.writeUInt32LE(0); // Flags
        body.writeUInt32LE(ropBuffer.length);
        body.writeBytes(ropBuffer);
        body.writeUInt32LE(256 * 1024); // MaxRopOut
        body.writeUInt32LE(0); // AuxiliaryBufferSize

        return mapiRequest(
            server.getApplication(),
            baseUrl,
            {
                Authorization: "jwt " + ownerToken,
                "X-RequestType": "Execute",
                "Content-Type": "application/mapi-http",
                ...(cookie ? { Cookie: cookie } : {}),
            },
            body.toBuffer(),
        );
    };

    beforeAll(async () => {
        await mongod.start();
        registerTestDoubles(objectFactory);
        await server.start();

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("mongo");
        if (conn instanceof MongoConnection) {
            mailboxRepo = conn.getMongoRepository("MailboxMongo");
            folderRepo = conn.getMongoRepository("FolderMongo");
            messageRepo = conn.getMongoRepository("MessageMongo");
            contactRepo = conn.getMongoRepository("ContactMongo");
            taskRepo = conn.getMongoRepository("TaskMongo");
        } else {
            throw new Error("Could not find mongo connection");
        }
        const aclConn: any = connMgr?.connections.get("acl");
        if (aclConn instanceof MongoConnection) {
            aclRepo = aclConn.getMongoRepository("AccessControlListMongo");
        } else {
            throw new Error("Could not find mongo acl connection");
        }
    });

    afterAll(async () => {
        await server.stop();
        await mongod.stop();
        await objectFactory.destroy();
    });

    beforeEach(async () => {
        for (const repo of [mailboxRepo, folderRepo, messageRepo, contactRepo, taskRepo, aclRepo]) {
            try {
                await repo.clear();
            } catch (err: any) {
                if (err.message !== "ns not found") {
                    throw err;
                }
            }
        }
    });

    describe("Connect", () => {
        it("Requires authentication.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { "X-RequestType": "Connect", "Content-Type": "application/mapi-http" },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(401);
        });

        it("Returns 404 when the caller owns no mailbox.", async () => {
            const result = await connect();
            expect(result.status).toBe(404);
        });

        it("Establishes a session, setting both MapiContext and MapiSequence cookies.", async () => {
            await createMailbox(owner.uid);
            const result = await connect();
            expect(result.status).toBe(200);
            const cookie = cookieHeaderFrom(result.headers["set-cookie"]);
            expect(cookie).toContain("MapiContext=");
            expect(cookie).toContain("MapiSequence=");
        });

        it("Returns a well-formed success response body (StatusCode/ErrorCode both zero) with the DisplayName echoed.", async () => {
            await createMailbox(owner.uid, { displayName: "Ada Lovelace" });
            const result = await connect();

            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            expect(reader.readUInt32LE()).toBe(60000); // PollsMax
            expect(reader.readUInt32LE()).toBe(3); // RetryCount
            expect(reader.readUInt32LE()).toBe(5000); // RetryDelay
            reader.readNullTerminatedString8(); // DnPrefix
            expect(reader.readNullTerminatedUtf16LE()).toBe("Ada Lovelace");
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
            expect(reader.hasMore()).toBe(false);
        });

        it("Sets the common MAPI/HTTP response headers.", async () => {
            await createMailbox(owner.uid);
            const result = await connect({ "X-RequestId": "req-123" });

            expect(result.headers["content-type"]).toContain("application/mapi-http");
            expect(result.headers["x-requesttype"]).toBe("Connect");
            expect(result.headers["x-requestid"]).toBe("req-123");
            expect(result.headers["x-responsecode"]).toBe("0");
        });
    });

    describe("Execute", () => {
        it("Returns a session-not-found error code when no MapiContext cookie is presented.", async () => {
            await createMailbox(owner.uid);
            const emptyRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const result = await execute("", emptyRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).not.toBe("0");
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).not.toBe(0); // ErrorCode
        });

        it("Echoes back an empty but well-formed ROP buffer, preserving the handle table, for a valid session.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const inputRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [7, 9] });
            const result = await execute(cookie, inputRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            expect(reader.readUInt32LE()).toBe(0); // StatusCode
            expect(reader.readUInt32LE()).toBe(0); // ErrorCode
            reader.readUInt32LE(); // Flags
            const responseRopBufferSize = reader.readUInt32LE();
            const responseRopBuffer = reader.readBytes(responseRopBufferSize);
            // RopSize field (2 bytes) + empty ropsList = 2, then the two preserved handle-table entries.
            expect(responseRopBuffer.readUInt16LE(0)).toBe(2);
            expect(responseRopBuffer.readUInt32LE(2)).toBe(7);
            expect(responseRopBuffer.readUInt32LE(6)).toBe(9);
            expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
        });
    });

    describe("RopLogon / RopRelease", () => {
        const buildRopLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe); // RopId
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01); // LogonFlags (Private)
            writer.writeUInt32LE(0); // OpenFlags
            writer.writeUInt32LE(0); // StoreState
            writer.writeUInt16LE(0); // EssdnSize
            return writer.toBuffer();
        };

        const buildRopReleaseRops = function (inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x01); // RopId
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            return writer.toBuffer();
        };

        it("Logs on to a real mailbox and returns 13 well-formed, distinct FIDs.", async () => {
            const mailbox = await createMailbox(owner.uid, { displayName: "Ada Lovelace" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const inputRop = encodeRopBuffer({ ropsList: buildRopLogonRops(0), handleTable: [0xffffffff] });
            const result = await execute(cookie, inputRop);

            expect(result.status).toBe(200);
            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            reader.readUInt32LE(); // StatusCode
            reader.readUInt32LE(); // ErrorCode
            reader.readUInt32LE(); // Flags
            const ropBufferSize = reader.readUInt32LE();
            const { ropsList, handleTable } = decodeRopBuffer(reader.readBytes(ropBufferSize));
            expect(handleTable).toEqual([0xffffffff]);

            const ropsReader = new BufferReader(ropsList);
            expect(ropsReader.readUInt8()).toBe(0xfe); // RopId
            expect(ropsReader.readUInt8()).toBe(0); // OutputHandleIndex
            expect(ropsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(ropsReader.readUInt8()).toBe(0x01); // LogonFlags, echoed

            const fids: bigint[] = [];
            for (let i = 0; i < 13; i++) {
                fids.push(ropsReader.readBigUInt64LE());
            }
            expect(new Set(fids).size).toBe(13); // all 13 FIDs are distinct

            ropsReader.readUInt8(); // ResponseFlags
            expect(decodeGuid(ropsReader)).toBe(mailbox.uid); // MailboxGuid
            expect(ropsReader.hasMore()).toBe(true); // ReplId/ReplGuid/LogonTime/GwartTime/StoreState follow
            // The FID<->Folder mapping itself (Inbox/Outbox/Sent/Deleted resolving to a real Folder when one
            // exists) is verified directly against session.folderIds in RopLogonHandler's own unit tests
            // (test/mapi/rop/RopLogonHandler.test.ts) - not observable from outside this black-box HTTP test,
            // since FIDs are opaque numbers to the client until a later RopOpenFolder (a future build step).
        });

        it("Releases a logon handle with no response bytes for that ROP, matching the spec's own captured example.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const logonRop = encodeRopBuffer({ ropsList: buildRopLogonRops(0), handleTable: [0xffffffff] });
            await execute(cookie, logonRop);

            const releaseRop = encodeRopBuffer({ ropsList: buildRopReleaseRops(0), handleTable: [0] });
            const result = await execute(cookie, releaseRop);

            expect(result.headers["x-responsecode"]).toBe("0");
            const reader = new BufferReader(result.body);
            reader.readUInt32LE(); // StatusCode
            reader.readUInt32LE(); // ErrorCode
            reader.readUInt32LE(); // Flags
            const ropBufferSize = reader.readUInt32LE();
            const { ropsList } = decodeRopBuffer(reader.readBytes(ropBufferSize));
            // RopRelease produces no response entry - the entire ropsList is empty.
            expect(ropsList.length).toBe(0);
        });
    });

    describe("Folder hierarchy browsing (RopOpenFolder / RopGetHierarchyTable / RopSetColumns / RopQueryRows)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetHierarchyTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x04);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        it("Browses the top-level folder list end to end: Logon, OpenFolder(root), GetHierarchyTable, SetColumns, QueryRows.", async () => {
            const mailbox = await createMailbox(owner.uid);
            await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            await createFolder(mailbox.uid, FolderType.USER, "Archive");
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the "root" FID.
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8(); // RopId
            logonRopsReader.readUInt8(); // OutputHandleIndex
            logonRopsReader.readUInt32LE(); // ReturnValue
            logonRopsReader.readUInt8(); // LogonFlags
            const rootFid = logonRopsReader.readBigUInt64LE(); // FolderIds[0] = Root

            // OpenFolder(root) at handle 1, using logon handle 0 as input.
            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, rootFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetHierarchyTable at handle 2, using folder handle 1 as input.
            const tableInput = encodeRopBuffer({ ropsList: buildGetHierarchyTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // SetColumns on table handle 2: DisplayName + ContentCount + ContentUnreadCount.
            const columns = [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString },
                { propertyId: 0x3602, propertyType: PropertyType.PtypInteger32 },
                { propertyId: 0x3603, propertyType: PropertyType.PtypInteger32 },
            ];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            const setColumnsResult = await execute(cookie, setColumnsInput);
            expect(setColumnsResult.headers["x-responsecode"]).toBe("0");

            // QueryRows on table handle 2.
            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8(); // RopId
            rowsReader.readUInt8(); // InputHandleIndex
            expect(rowsReader.readUInt32LE()).toBe(0); // ReturnValue
            rowsReader.readUInt8(); // Origin
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const names: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                names.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
                readPropertyValue(rowsReader, PropertyType.PtypInteger32); // ContentCount
                readPropertyValue(rowsReader, PropertyType.PtypInteger32); // ContentUnreadCount
            }
            expect(names.sort()).toEqual(["Archive", "Inbox"]);
            expect(rowsReader.hasMore()).toBe(false);
        });
    });

    describe("Message listing (RopGetContentsTable / RopSetColumns / RopQueryRows)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetContentsTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x05);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        it("Lists a folder's messages end to end: Logon, OpenFolder(inbox), GetContentsTable, SetColumns, QueryRows.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            await createMessage(mailbox.uid, inbox.uid, { subject: "Hello World" });
            await createMessage(mailbox.uid, inbox.uid, { subject: "Second Message" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID.
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8(); // RopId
            logonRopsReader.readUInt8(); // OutputHandleIndex
            logonRopsReader.readUInt32LE(); // ReturnValue
            logonRopsReader.readUInt8(); // LogonFlags
            logonRopsReader.readBigUInt64LE(); // FolderIds[0] = Root
            logonRopsReader.readBigUInt64LE(); // FolderIds[1] = Deferred Action
            logonRopsReader.readBigUInt64LE(); // FolderIds[2] = Spooler Queue
            logonRopsReader.readBigUInt64LE(); // FolderIds[3] = IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // FolderIds[4] = Inbox

            // OpenFolder(inbox) at handle 1, using logon handle 0 as input.
            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetContentsTable at handle 2, using folder handle 1 as input.
            const tableInput = encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // SetColumns on table handle 2: Subject.
            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            const setColumnsResult = await execute(cookie, setColumnsInput);
            expect(setColumnsResult.headers["x-responsecode"]).toBe("0");

            // QueryRows on table handle 2.
            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8(); // RopId
            rowsReader.readUInt8(); // InputHandleIndex
            expect(rowsReader.readUInt32LE()).toBe(0); // ReturnValue
            rowsReader.readUInt8(); // Origin
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const subjects: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                subjects.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(subjects.sort()).toEqual(["Hello World", "Second Message"]);
            expect(rowsReader.hasMore()).toBe(false);
        });
    });

    describe("Contacts/Tasks folder browsing (RopGetContentsTable over CONTACTS/TASKS folders)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetTableRops = function (ropId: number, inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(ropId);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0); // TableFlags - Standard
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SetColumnsFlags
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // QueryRowsFlags
            writer.writeUInt8(1); // ForwardRead
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        /** Discovers `wellKnownFolderName`'s real FID via the root folder's own hierarchy table - Contacts/
         * Tasks aren't part of RopLogon's fixed 13-folder `FolderIds` list (see RopLogonHandler's own doc
         * comment), so a real client (and this test) finds them the same way it finds any other top-level
         * folder: OpenFolder(root) -> GetHierarchyTable -> SetColumns([DisplayName, FolderId]) -> QueryRows. */
        async function discoverFolderId(cookie: string, wellKnownFolderName: string): Promise<bigint> {
            const logonInput = encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] });
            const logonResult = await execute(cookie, logonInput);
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const logonRopBufferSize = logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonRopBufferSize));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            const rootFid = logonRopsReader.readBigUInt64LE();

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, rootFid), handleTable: [0xffffffff, 0xffffffff] });
            await execute(cookie, openFolderInput);

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x04, 1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            await execute(cookie, tableInput);

            const columns = [
                { propertyId: 0x3001, propertyType: PropertyType.PtypString }, // DisplayName
                { propertyId: 0x6748, propertyType: PropertyType.PtypInteger64 }, // FolderId
            ];
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            rowsReader.readUInt32LE();
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8(); // PropertyRow Flags
                const name = readPropertyValue(rowsReader, PropertyType.PtypString) as string;
                const fid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;
                if (name === wellKnownFolderName) {
                    return fid;
                }
            }
            throw new Error(`Folder named "${wellKnownFolderName}" was not found in the root hierarchy table.`);
        }

        it("Lists a CONTACTS folder's contacts end to end, resolved by displayName via ContactTarget.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const contactsFolder = await createFolder(mailbox.uid, FolderType.CONTACTS, "Contacts");
            await createContact(mailbox.uid, contactsFolder.uid, { displayName: "Jane Doe" });
            await createContact(mailbox.uid, contactsFolder.uid, { displayName: "John Smith" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const contactsFid = await discoverFolderId(cookie, "Contacts");

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 3, contactsFid), handleTable: [0xffffffff, 0xffffffff] });
            const openFolderResult = await execute(cookie, openFolderInput);
            const openFolderReader = new BufferReader(openFolderResult.body);
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            openFolderReader.readUInt32LE();
            const openFolderRopBufferSize = openFolderReader.readUInt32LE();
            const { ropsList: openFolderRopsList } = decodeRopBuffer(openFolderReader.readBytes(openFolderRopBufferSize));
            const openFolderRopsReader = new BufferReader(openFolderRopsList);
            openFolderRopsReader.readUInt8();
            openFolderRopsReader.readUInt8();
            expect(openFolderRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x05, 3, 4), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            const tableResult = await execute(cookie, tableInput);
            const tableReader = new BufferReader(tableResult.body);
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            tableReader.readUInt32LE();
            const tableRopBufferSize = tableReader.readUInt32LE();
            const { ropsList: tableRopsList } = decodeRopBuffer(tableReader.readBytes(tableRopBufferSize));
            const tableRopsReader = new BufferReader(tableRopsList);
            tableRopsReader.readUInt8();
            tableRopsReader.readUInt8();
            expect(tableRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]; // PidTagSubject -> displayName
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(4, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(4, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt32LE()).toBe(0);
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const names: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8();
                names.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(names.sort()).toEqual(["Jane Doe", "John Smith"]);
        });

        it("Lists a TASKS folder's tasks end to end, resolved by title via TaskTarget.", async () => {
            const mailbox = await createMailbox(owner.uid);
            const tasksFolder = await createFolder(mailbox.uid, FolderType.TASKS, "Tasks");
            await createTask(mailbox.uid, tasksFolder.uid, { title: "Ship it" });
            await createTask(mailbox.uid, tasksFolder.uid, { title: "Write docs" });
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const tasksFid = await discoverFolderId(cookie, "Tasks");

            const openFolderInput = encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 3, tasksFid), handleTable: [0xffffffff, 0xffffffff] });
            await execute(cookie, openFolderInput);

            const tableInput = encodeRopBuffer({ ropsList: buildGetTableRops(0x05, 3, 4), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] });
            await execute(cookie, tableInput);

            const columns = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]; // PidTagSubject -> title
            const setColumnsInput = encodeRopBuffer({ ropsList: buildSetColumnsRops(4, columns), handleTable: [0xffffffff] });
            await execute(cookie, setColumnsInput);

            const queryRowsInput = encodeRopBuffer({ ropsList: buildQueryRowsRops(4, 10), handleTable: [0xffffffff] });
            const queryRowsResult = await execute(cookie, queryRowsInput);
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const queryRowsBufferSize = queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsBufferSize));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt32LE()).toBe(0);
            rowsReader.readUInt8();
            const rowCount = rowsReader.readUInt16LE();
            expect(rowCount).toBe(2);

            const titles: string[] = [];
            for (let i = 0; i < rowCount; i++) {
                rowsReader.readUInt8();
                titles.push(readPropertyValue(rowsReader, PropertyType.PtypString) as string);
            }
            expect(titles.sort()).toEqual(["Ship it", "Write docs"]);
        });
    });

    describe("Message reading (RopOpenMessage / RopGetPropertiesSpecific / RopOpenStream / RopReadStream)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildOpenFolderRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x02);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            writer.writeBigUInt64LE(folderId);
            return writer.toBuffer();
        };

        const buildGetContentsTableRops = function (inputHandleIndex: number, outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x05);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0);
            return writer.toBuffer();
        };

        const buildSetColumnsRops = function (inputHandleIndex: number, tags: { propertyId: number; propertyType: PropertyType }[]): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x12);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildQueryRowsRops = function (inputHandleIndex: number, rowCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x15);
            writer.writeUInt8(0);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0);
            writer.writeUInt8(1);
            writer.writeUInt16LE(rowCount);
            return writer.toBuffer();
        };

        const buildOpenMessageRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            folderId: bigint,
            messageId: bigint,
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x03);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt16LE(0); // CodePageId
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0); // OpenModeFlags
            writer.writeBigUInt64LE(messageId);
            return writer.toBuffer();
        };

        const buildGetPropertiesSpecificRops = function (
            inputHandleIndex: number,
            tags: { propertyId: number; propertyType: PropertyType }[],
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x07);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(0); // PropertySizeLimit
            writer.writeUInt16LE(0); // WantUnicode
            writer.writeUInt16LE(tags.length);
            for (const tag of tags) {
                writePropertyTag(writer, tag);
            }
            return writer.toBuffer();
        };

        const buildOpenStreamRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            propertyTag: { propertyId: number; propertyType: PropertyType },
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2b);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writePropertyTag(writer, propertyTag);
            writer.writeUInt8(0); // OpenModeFlags
            return writer.toBuffer();
        };

        const buildReadStreamRops = function (inputHandleIndex: number, byteCount: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2c);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(byteCount);
            return writer.toBuffer();
        };

        it("Reads a real message end to end: OpenMessage, GetPropertiesSpecific(Subject), OpenStream+ReadStream(Body).", async () => {
            const mailbox = await createMailbox(owner.uid);
            const inbox = await createFolder(mailbox.uid, FolderType.INBOX, "Inbox");
            const message = await createMessage(mailbox.uid, inbox.uid, { subject: "Read Me" });
            await blobStore().put(
                message.bodyBlobKey,
                Buffer.from("From: sender@example.com\r\nTo: owner@example.com\r\nSubject: Read Me\r\n\r\nActual body text."),
            );
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID.
            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            logonRopsReader.readBigUInt64LE(); // Root
            logonRopsReader.readBigUInt64LE(); // Deferred Action
            logonRopsReader.readBigUInt64LE(); // Spooler Queue
            logonRopsReader.readBigUInt64LE(); // IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // Inbox

            // OpenFolder(inbox) at handle 1.
            await execute(cookie, encodeRopBuffer({ ropsList: buildOpenFolderRops(0, 1, inboxFid), handleTable: [0xffffffff, 0xffffffff] }));

            // GetContentsTable at handle 2, SetColumns(Subject, Mid), QueryRows - to learn the message's MID.
            await execute(cookie, encodeRopBuffer({ ropsList: buildGetContentsTableRops(1, 2), handleTable: [0xffffffff, 0xffffffff, 0xffffffff] }));
            const columns = [
                { propertyId: 0x0037, propertyType: PropertyType.PtypString },
                { propertyId: 0x674a, propertyType: PropertyType.PtypInteger64 },
            ];
            await execute(cookie, encodeRopBuffer({ ropsList: buildSetColumnsRops(2, columns), handleTable: [0xffffffff] }));
            const queryRowsResult = await execute(cookie, encodeRopBuffer({ ropsList: buildQueryRowsRops(2, 10), handleTable: [0xffffffff] }));
            const queryRowsReader = new BufferReader(queryRowsResult.body);
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            queryRowsReader.readUInt32LE();
            const { ropsList: queryRowsList } = decodeRopBuffer(queryRowsReader.readBytes(queryRowsReader.readUInt32LE()));
            const rowsReader = new BufferReader(queryRowsList);
            rowsReader.readUInt8();
            rowsReader.readUInt8();
            rowsReader.readUInt32LE();
            rowsReader.readUInt8();
            expect(rowsReader.readUInt16LE()).toBe(1); // RowCount
            rowsReader.readUInt8(); // PropertyRow Flags
            expect(readPropertyValue(rowsReader, PropertyType.PtypString)).toBe("Read Me");
            const mid = readPropertyValue(rowsReader, PropertyType.PtypInteger64) as bigint;

            // OpenMessage at handle 3, using the MID just learned.
            const openMessageResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildOpenMessageRops(1, 3, inboxFid, mid), handleTable: [0xffffffff, 0xffffffff] }),
            );
            const openMessageReader = new BufferReader(openMessageResult.body);
            openMessageReader.readUInt32LE();
            openMessageReader.readUInt32LE();
            openMessageReader.readUInt32LE();
            const { ropsList: openMessageRopsList } = decodeRopBuffer(openMessageReader.readBytes(openMessageReader.readUInt32LE()));
            const openMessageRopsReader = new BufferReader(openMessageRopsList);
            openMessageRopsReader.readUInt8();
            openMessageRopsReader.readUInt8();
            expect(openMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // GetPropertiesSpecific(Subject) on message handle 3.
            const getPropsInput = encodeRopBuffer({
                ropsList: buildGetPropertiesSpecificRops(3, [{ propertyId: 0x0037, propertyType: PropertyType.PtypString }]),
                handleTable: [0xffffffff],
            });
            const getPropsResult = await execute(cookie, getPropsInput);
            const getPropsReader = new BufferReader(getPropsResult.body);
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            getPropsReader.readUInt32LE();
            const { ropsList: getPropsRopsList } = decodeRopBuffer(getPropsReader.readBytes(getPropsReader.readUInt32LE()));
            const getPropsRopsReader = new BufferReader(getPropsRopsList);
            getPropsRopsReader.readUInt8();
            getPropsRopsReader.readUInt8();
            expect(getPropsRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            getPropsRopsReader.readUInt8(); // PropertyRow Flags
            expect(readPropertyValue(getPropsRopsReader, PropertyType.PtypString)).toBe("Read Me");

            // OpenStream(PidTagBody) at handle 4, then ReadStream to fetch the actual body text.
            const openStreamInput = encodeRopBuffer({
                ropsList: buildOpenStreamRops(3, 4, { propertyId: 0x1000, propertyType: PropertyType.PtypString }),
                handleTable: [0xffffffff, 0xffffffff],
            });
            const openStreamResult = await execute(cookie, openStreamInput);
            const openStreamReader = new BufferReader(openStreamResult.body);
            openStreamReader.readUInt32LE();
            openStreamReader.readUInt32LE();
            openStreamReader.readUInt32LE();
            const { ropsList: openStreamRopsList } = decodeRopBuffer(openStreamReader.readBytes(openStreamReader.readUInt32LE()));
            const openStreamRopsReader = new BufferReader(openStreamRopsList);
            openStreamRopsReader.readUInt8();
            openStreamRopsReader.readUInt8();
            expect(openStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            const streamSize = openStreamRopsReader.readUInt32LE();
            expect(streamSize).toBeGreaterThan(0);

            const readStreamInput = encodeRopBuffer({ ropsList: buildReadStreamRops(4, 1000), handleTable: [0xffffffff] });
            const readStreamResult = await execute(cookie, readStreamInput);
            const readStreamReader = new BufferReader(readStreamResult.body);
            readStreamReader.readUInt32LE();
            readStreamReader.readUInt32LE();
            readStreamReader.readUInt32LE();
            const { ropsList: readStreamRopsList } = decodeRopBuffer(readStreamReader.readBytes(readStreamReader.readUInt32LE()));
            const readStreamRopsReader = new BufferReader(readStreamRopsList);
            readStreamRopsReader.readUInt8();
            readStreamRopsReader.readUInt8();
            expect(readStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            const dataSize = readStreamRopsReader.readUInt16LE();
            expect(dataSize).toBe(streamSize);
            expect(readPropertyValue(readStreamRopsReader, PropertyType.PtypString)).toBe("Actual body text.");
        });
    });

    describe("Compose/Send (RopCreateMessage / RopSetProperties / RopOpenStream / RopWriteStream / RopSaveChangesMessage / RopSubmitMessage)", () => {
        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildCreateMessageRops = function (inputHandleIndex: number, outputHandleIndex: number, folderId: bigint): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x06);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt16LE(0); // CodePageId
            writer.writeBigUInt64LE(folderId);
            writer.writeUInt8(0); // AssociatedFlag
            return writer.toBuffer();
        };

        const buildSetPropertiesRops = function (inputHandleIndex: number, values: TaggedPropertyValue[]): Buffer {
            const valuesWriter = new BufferWriter();
            for (const value of values) {
                writeTaggedPropertyValue(valuesWriter, value);
            }
            const valuesBytes = valuesWriter.toBuffer();
            const writer = new BufferWriter();
            writer.writeUInt8(0x0a);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(2 + valuesBytes.length);
            writer.writeUInt16LE(values.length);
            writer.writeBytes(valuesBytes);
            return writer.toBuffer();
        };

        const buildOpenStreamRops = function (
            inputHandleIndex: number,
            outputHandleIndex: number,
            propertyTag: { propertyId: number; propertyType: PropertyType },
            openModeFlags: number,
        ): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2b);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(outputHandleIndex);
            writePropertyTag(writer, propertyTag);
            writer.writeUInt8(openModeFlags);
            return writer.toBuffer();
        };

        const buildWriteStreamRops = function (inputHandleIndex: number, data: Buffer): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x2d);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt16LE(data.length);
            writer.writeBytes(data);
            return writer.toBuffer();
        };

        const buildSaveChangesMessageRops = function (responseHandleIndex: number, inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x0c);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(responseHandleIndex);
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SaveFlags
            return writer.toBuffer();
        };

        const buildSubmitMessageRops = function (inputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0x32);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0); // SubmitFlags
            return writer.toBuffer();
        };

        it("Composes and sends a real message end to end, saving a Sent Items copy.", async () => {
            mailTransport().sent = [];
            const mailbox = await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - discover the Inbox FID (used only as RopCreateMessage's target folder; the
            // actual saved copy always lands in Sent Items regardless, matching RopSubmitMessageHandler's own
            // documented behavior).
            const logonResult = await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));
            const logonReader = new BufferReader(logonResult.body);
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            logonReader.readUInt32LE();
            const { ropsList: logonRopsList } = decodeRopBuffer(logonReader.readBytes(logonReader.readUInt32LE()));
            const logonRopsReader = new BufferReader(logonRopsList);
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt8();
            logonRopsReader.readUInt32LE();
            logonRopsReader.readUInt8();
            logonRopsReader.readBigUInt64LE(); // Root
            logonRopsReader.readBigUInt64LE(); // Deferred Action
            logonRopsReader.readBigUInt64LE(); // Spooler Queue
            logonRopsReader.readBigUInt64LE(); // IPM Subtree
            const inboxFid = logonRopsReader.readBigUInt64LE(); // Inbox

            // RopCreateMessage at handle 3.
            const createMessageResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildCreateMessageRops(0, 3, inboxFid), handleTable: [0xffffffff, 0xffffffff] }),
            );
            const createMessageReader = new BufferReader(createMessageResult.body);
            createMessageReader.readUInt32LE();
            createMessageReader.readUInt32LE();
            createMessageReader.readUInt32LE();
            const { ropsList: createMessageRopsList } = decodeRopBuffer(createMessageReader.readBytes(createMessageReader.readUInt32LE()));
            const createMessageRopsReader = new BufferReader(createMessageRopsList);
            createMessageRopsReader.readUInt8();
            createMessageRopsReader.readUInt8();
            expect(createMessageRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            // RopSetProperties(Subject, DisplayTo) on message handle 3.
            const setPropertiesInput = encodeRopBuffer({
                ropsList: buildSetPropertiesRops(3, [
                    { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hello From MAPI" },
                    { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: "recipient@example.com" },
                ]),
                handleTable: [0xffffffff],
            });
            const setPropertiesResult = await execute(cookie, setPropertiesInput);
            expect(setPropertiesResult.headers["x-responsecode"]).toBe("0");

            // RopOpenStream(PidTagBody, Create) at handle 4, then RopWriteStream to set the body text.
            const openStreamInput = encodeRopBuffer({
                ropsList: buildOpenStreamRops(3, 4, { propertyId: 0x1000, propertyType: PropertyType.PtypString }, 0x02),
                handleTable: [0xffffffff, 0xffffffff],
            });
            await execute(cookie, openStreamInput);

            const bodyText = "This message was composed entirely via MAPI ROPs.";
            const bodyBytes = Buffer.concat([Buffer.from(bodyText, "utf16le"), Buffer.from([0, 0])]);
            const writeStreamInput = encodeRopBuffer({ ropsList: buildWriteStreamRops(4, bodyBytes), handleTable: [0xffffffff] });
            const writeStreamResult = await execute(cookie, writeStreamInput);
            const writeStreamReader = new BufferReader(writeStreamResult.body);
            writeStreamReader.readUInt32LE();
            writeStreamReader.readUInt32LE();
            writeStreamReader.readUInt32LE();
            const { ropsList: writeStreamRopsList } = decodeRopBuffer(writeStreamReader.readBytes(writeStreamReader.readUInt32LE()));
            const writeStreamRopsReader = new BufferReader(writeStreamRopsList);
            writeStreamRopsReader.readUInt8();
            writeStreamRopsReader.readUInt8();
            expect(writeStreamRopsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(writeStreamRopsReader.readUInt16LE()).toBe(bodyBytes.length); // WrittenSize

            // RopSaveChangesMessage at response handle 7, on message handle 3.
            const saveChangesInput = encodeRopBuffer({ ropsList: buildSaveChangesMessageRops(7, 3), handleTable: [0xffffffff] });
            const saveChangesResult = await execute(cookie, saveChangesInput);
            const saveChangesReader = new BufferReader(saveChangesResult.body);
            saveChangesReader.readUInt32LE();
            saveChangesReader.readUInt32LE();
            saveChangesReader.readUInt32LE();
            const { ropsList: saveChangesRopsList } = decodeRopBuffer(saveChangesReader.readBytes(saveChangesReader.readUInt32LE()));
            const saveChangesRopsReader = new BufferReader(saveChangesRopsList);
            saveChangesRopsReader.readUInt8();
            saveChangesRopsReader.readUInt8();
            expect(saveChangesRopsReader.readUInt32LE()).toBe(0); // ReturnValue

            // RopSubmitMessage on message handle 3 - the real send trigger.
            const submitInput = encodeRopBuffer({ ropsList: buildSubmitMessageRops(3), handleTable: [0xffffffff] });
            const submitResult = await execute(cookie, submitInput);
            const submitReader = new BufferReader(submitResult.body);
            submitReader.readUInt32LE();
            submitReader.readUInt32LE();
            submitReader.readUInt32LE();
            const { ropsList: submitRopsList } = decodeRopBuffer(submitReader.readBytes(submitReader.readUInt32LE()));
            const submitRopsReader = new BufferReader(submitRopsList);
            submitRopsReader.readUInt8();
            submitRopsReader.readUInt8();
            expect(submitRopsReader.readUInt32LE()).toBe(0); // ReturnValue - success

            expect(mailTransport().sent.length).toBe(1);
            expect(mailTransport().sent[0].envelopeFrom).toBe(mailbox.primarySmtpAddress);
            expect(mailTransport().sent[0].envelopeTo).toEqual(["recipient@example.com"]);
            expect(mailTransport().sent[0].raw.toString("utf-8")).toContain("Hello From MAPI");

            const sentMessages = await messageRepo.find({ subject: "Hello From MAPI" }).toArray();
            expect(sentMessages.length).toBe(1);
            expect(sentMessages[0].recipients).toEqual([{ address: "recipient@example.com", type: RecipientType.TO }]);
            const savedRaw = await blobStore().get(sentMessages[0].bodyBlobKey);
            expect(savedRaw.toString("utf-8")).toContain(bodyText);
        });
    });

    describe("RopGetPropertyIdsFromNames", () => {
        const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";

        const buildLogonRops = function (outputHandleIndex: number): Buffer {
            const writer = new BufferWriter();
            writer.writeUInt8(0xfe);
            writer.writeUInt8(0);
            writer.writeUInt8(outputHandleIndex);
            writer.writeUInt8(0x01);
            writer.writeUInt32LE(0);
            writer.writeUInt32LE(0);
            writer.writeUInt16LE(0);
            return writer.toBuffer();
        };

        const buildGetPropertyIdsFromNamesRops = function (
            inputHandleIndex: number,
            propertyNames: { guid: string; lid: number }[],
        ): Buffer {
            const namesWriter = new BufferWriter();
            for (const name of propertyNames) {
                namesWriter.writeUInt8(0x00); // Kind - LID
                namesWriter.writeBytes(encodeGuid(name.guid));
                namesWriter.writeUInt32LE(name.lid);
            }
            const namesBytes = namesWriter.toBuffer();

            const writer = new BufferWriter();
            writer.writeUInt8(0x56);
            writer.writeUInt8(0); // LogonId
            writer.writeUInt8(inputHandleIndex);
            writer.writeUInt8(0x02); // Flags - assign new IDs
            writer.writeUInt16LE(propertyNames.length);
            writer.writeBytes(namesBytes);
            return writer.toBuffer();
        };

        it("Resolves named properties into stable, session-scoped numeric IDs starting at 0x8000.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            // Logon (handle 0) - any real handle works as RopGetPropertyIdsFromNames's InputHandleIndex.
            await execute(cookie, encodeRopBuffer({ ropsList: buildLogonRops(0), handleTable: [0xffffffff] }));

            // Arbitrary LIDs within PSETID_Appointment - this test only exercises the ROP's own mechanics
            // (resolving a PropertyName to a stable session-scoped numeric ID), not any specific real Calendar
            // property's exact LID value (those are pinned against MS-OXPROPS individually where they're
            // actually consumed - see PropertyResolvers.ts's own Calendar property mapping, a later step).
            const propertyNames = [
                { guid: PSETID_APPOINTMENT, lid: 0x8208 },
                { guid: PSETID_APPOINTMENT, lid: 0x820d },
            ];
            const result = await execute(cookie, encodeRopBuffer({ ropsList: buildGetPropertyIdsFromNamesRops(0, propertyNames), handleTable: [0xffffffff] }));

            const reader = new BufferReader(result.body);
            reader.readUInt32LE();
            reader.readUInt32LE();
            reader.readUInt32LE();
            const { ropsList } = decodeRopBuffer(reader.readBytes(reader.readUInt32LE()));
            const ropsReader = new BufferReader(ropsList);
            expect(ropsReader.readUInt8()).toBe(0x56); // RopId
            expect(ropsReader.readUInt8()).toBe(0); // InputHandleIndex, echoed
            expect(ropsReader.readUInt32LE()).toBe(0); // ReturnValue
            expect(ropsReader.readUInt16LE()).toBe(2); // PropertyIdCount
            const id1 = ropsReader.readUInt16LE();
            const id2 = ropsReader.readUInt16LE();
            expect(id1).toBe(0x8000);
            expect(id2).toBe(0x8001);
            expect(ropsReader.hasMore()).toBe(false);

            // Resolving the exact same PropertyName again in a later Execute call reuses the same ID - the
            // registry is session-scoped, not re-assigned per call.
            const secondResult = await execute(
                cookie,
                encodeRopBuffer({ ropsList: buildGetPropertyIdsFromNamesRops(0, [propertyNames[0]]), handleTable: [0xffffffff] }),
            );
            const secondReader = new BufferReader(secondResult.body);
            secondReader.readUInt32LE();
            secondReader.readUInt32LE();
            secondReader.readUInt32LE();
            const { ropsList: secondRopsList } = decodeRopBuffer(secondReader.readBytes(secondReader.readUInt32LE()));
            const secondRopsReader = new BufferReader(secondRopsList);
            secondRopsReader.readUInt8();
            secondRopsReader.readUInt8();
            secondRopsReader.readUInt32LE();
            secondRopsReader.readUInt16LE();
            expect(secondRopsReader.readUInt16LE()).toBe(id1);
        });
    });

    describe("Disconnect", () => {
        it("Releases the session such that a subsequent Execute reports session-not-found.", async () => {
            await createMailbox(owner.uid);
            const connectResult = await connect();
            const cookie = cookieHeaderFrom(connectResult.headers["set-cookie"]);

            const disconnectResult = await mapiRequest(
                server.getApplication(),
                baseUrl,
                {
                    Authorization: "jwt " + ownerToken,
                    "X-RequestType": "Disconnect",
                    "Content-Type": "application/mapi-http",
                    Cookie: cookie,
                },
                Buffer.alloc(0),
            );
            expect(disconnectResult.status).toBe(200);

            const emptyRop = encodeRopBuffer({ ropsList: Buffer.alloc(0), handleTable: [] });
            const executeResult = await execute(cookie, emptyRop);
            expect(executeResult.headers["x-responsecode"]).not.toBe("0");
        });
    });

    describe("Unrecognized/deferred request types", () => {
        it("Returns 400 for a missing/unrecognized X-RequestType.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                { Authorization: "jwt " + ownerToken, "Content-Type": "application/mapi-http" },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(400);
        });

        it("Returns 501 for the recognized-but-deferred NotificationWait request type.", async () => {
            const result = await mapiRequest(
                server.getApplication(),
                baseUrl,
                {
                    Authorization: "jwt " + ownerToken,
                    "X-RequestType": "NotificationWait",
                    "Content-Type": "application/mapi-http",
                },
                Buffer.alloc(0),
            );
            expect(result.status).toBe(501);
        });
    });
});
