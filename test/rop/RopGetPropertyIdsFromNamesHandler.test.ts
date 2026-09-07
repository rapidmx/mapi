///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { encodeGuid } from "../../src/codec/MapiGuid.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { RopGetPropertyIdsFromNamesHandler } from "../../src/rop/RopGetPropertyIdsFromNamesHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";

type PropertyNameFixture =
    | { kind: 0x00; guid: string; lid: number }
    | { kind: 0x01; guid: string; name: string }
    | { kind: 0xff; guid: string };

function writePropertyNameFixture(writer: BufferWriter, fixture: PropertyNameFixture): void {
    writer.writeUInt8(fixture.kind);
    writer.writeBytes(encodeGuid(fixture.guid));
    if (fixture.kind === 0x00) {
        writer.writeUInt32LE(fixture.lid);
    } else if (fixture.kind === 0x01) {
        const nameBytes = Buffer.concat([Buffer.from(fixture.name, "utf16le"), Buffer.from([0x00, 0x00])]);
        writer.writeUInt8(nameBytes.length);
        writer.writeBytes(nameBytes);
    }
}

function buildRequest({
    logonId = 0,
    inputHandleIndex = 5,
    flags = 0x02,
    propertyNames = [] as PropertyNameFixture[],
}): Buffer {
    const namesWriter = new BufferWriter();
    for (const fixture of propertyNames) {
        writePropertyNameFixture(namesWriter, fixture);
    }
    const namesBytes = namesWriter.toBuffer();

    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt8(flags);
    writer.writeUInt16LE(propertyNames.length);
    writer.writeBytes(namesBytes);
    return writer.toBuffer();
}

function makeContext(): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopGetPropertyIdsFromNamesHandler Tests", () => {
    it("Has RopId 0x56.", () => {
        expect(new RopGetPropertyIdsFromNamesHandler().ropId).toBe(0x56);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex references no handle at all.", () => {
        const context = makeContext();
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Assigns numeric IDs (starting at 0x8000) for Kind=LID property names, in request order.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        const request = buildRequest({
            propertyNames: [
                { kind: 0x00, guid: PSETID_APPOINTMENT, lid: 0x8208 },
                { kind: 0x00, guid: PSETID_APPOINTMENT, lid: 0x8501 },
            ],
        });
        handler.handle(new BufferReader(request), writer, context);

        const response = new BufferReader(writer.toBuffer());
        expect(response.readUInt8()).toBe(0x56);
        expect(response.readUInt8()).toBe(5);
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt16LE()).toBe(2); // PropertyIdCount
        const id1 = response.readUInt16LE();
        const id2 = response.readUInt16LE();
        expect(id1).toBe(0x8000);
        expect(id2).toBe(0x8001);
        expect(response.hasMore()).toBe(false);
    });

    it("Reuses an already-assigned ID for a repeated PropertyName within the same session.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const existingId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: 0x8208 });
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        const request = buildRequest({ propertyNames: [{ kind: 0x00, guid: PSETID_APPOINTMENT, lid: 0x8208 }] });
        handler.handle(new BufferReader(request), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        response.readUInt16LE();
        expect(response.readUInt16LE()).toBe(existingId);
    });

    it("Resolves a Kind=Name property name too.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        const request = buildRequest({ propertyNames: [{ kind: 0x01, guid: PSETID_APPOINTMENT, name: "SomeProp" }] });
        handler.handle(new BufferReader(request), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(1);
        expect(response.readUInt16LE()).toBe(0x8000);
    });

    it("Resolves a Kind=Name property name whose NameSize excludes the null terminator too.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        // Hand-built request bypassing writePropertyNameFixture's own (terminator-inclusive) NameSize
        // convention, to exercise the case where NameSize covers only the content bytes.
        const nameContentBytes = Buffer.from("SomeProp", "utf16le");
        const namesWriter = new BufferWriter();
        namesWriter.writeUInt8(0x01); // Kind - Name
        namesWriter.writeBytes(encodeGuid(PSETID_APPOINTMENT));
        namesWriter.writeUInt8(nameContentBytes.length);
        namesWriter.writeBytes(nameContentBytes);
        const namesBytes = namesWriter.toBuffer();

        const requestWriter = new BufferWriter();
        requestWriter.writeUInt8(0); // LogonId
        requestWriter.writeUInt8(5); // InputHandleIndex
        requestWriter.writeUInt8(0x02); // Flags
        requestWriter.writeUInt16LE(1);
        requestWriter.writeBytes(namesBytes);

        handler.handle(new BufferReader(requestWriter.toBuffer()), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(1);
        expect(response.readUInt16LE()).toBe(0x8000);
    });

    it("Maps a Kind=0xFF (no PropertyName) entry to 0x0000, per spec.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        const request = buildRequest({ propertyNames: [{ kind: 0xff, guid: PSETID_APPOINTMENT }] });
        handler.handle(new BufferReader(request), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        response.readUInt32LE();
        expect(response.readUInt16LE()).toBe(1);
        expect(response.readUInt16LE()).toBe(0x0000);
    });

    it("Returns an empty PropertyIds array for a request with zero PropertyNames.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "logon", entityUid: "mailbox-1" };
        const handler = new RopGetPropertyIdsFromNamesHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(response.readUInt16LE()).toBe(0);
        expect(response.hasMore()).toBe(false);
    });
});
