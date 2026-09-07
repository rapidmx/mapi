///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, writeTaggedPropertyValue, type TaggedPropertyValue } from "../../src/codec/PropertyValue.js";
import { assignOrGetNamedPropertyId } from "../../src/rop/NamedPropertyRegistry.js";
import { RopSetPropertiesHandler } from "../../src/rop/RopSetPropertiesHandler.js";
import type { RopContext } from "../../src/rop/RopHandler.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

const PSETID_APPOINTMENT = "00062002-0000-0000-c000-000000000046";
const PSETID_COMMON = "00062008-0000-0000-c000-000000000046";
const PSETID_MEETING = "6ed8da90-450b-101b-98da-00aa003f1305";
const LID_LOCATION = 0x8208;
const LID_APPOINTMENT_START_WHOLE = 0x820d;
const LID_APPOINTMENT_RECUR = 0x8216;
const LID_REMINDER_DELTA = 0x8501;
const LID_GLOBAL_OBJECT_ID = 0x00000003;

function buildRequest({ logonId = 0, inputHandleIndex = 5, values = [] as TaggedPropertyValue[] }): Buffer {
    const valuesWriter = new BufferWriter();
    for (const value of values) {
        writeTaggedPropertyValue(valuesWriter, value);
    }
    const valuesBytes = valuesWriter.toBuffer();

    const writer = new BufferWriter();
    writer.writeUInt8(logonId);
    writer.writeUInt8(inputHandleIndex);
    writer.writeUInt16LE(2 + valuesBytes.length); // PropertyValueSize (covers PropertyValueCount + PropertyValues)
    writer.writeUInt16LE(values.length);
    writer.writeBytes(valuesBytes);
    return writer.toBuffer();
}

function makeContext(): RopContext {
    return {
        mailboxUid: "mailbox-1",
        userUid: "user-1",
        session: new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" }),
        folderRepo: {} as any,
        messageRepo: {} as any,
        calendarEventRepo: {} as any,
        mailboxRepo: {} as any,
        folderClass: {} as any,
        messageClass: {} as any,
        calendarEventClass: {} as any,
        scanPipeline: {} as any,
        mailTransport: {} as any,
        blobStore: {} as any,
    };
}

describe("RopSetPropertiesHandler Tests", () => {
    it("Has RopId 0x0A.", () => {
        expect(new RopSetPropertiesHandler().ropId).toBe(0x0a);
    });

    it("Returns MAPI_E_INVALID_OBJECT when InputHandleIndex isn't a message handle.", () => {
        const context = makeContext();
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        handler.handle(new BufferReader(buildRequest({})), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0x80070005);
        expect(response.hasMore()).toBe(false);
    });

    it("Tracks Subject/DisplayTo/DisplayCc/DisplayBcc/Body onto the message handle's draftProperties.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1", draftProperties: {} };
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        const values: TaggedPropertyValue[] = [
            { propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hello" },
            { propertyId: 0x0e04, propertyType: PropertyType.PtypString, value: "to@example.com" },
            { propertyId: 0x0e03, propertyType: PropertyType.PtypString, value: "cc@example.com" },
            { propertyId: 0x0e02, propertyType: PropertyType.PtypString, value: "bcc@example.com" },
            { propertyId: 0x1000, propertyType: PropertyType.PtypString, value: "Body text" },
        ];
        handler.handle(new BufferReader(buildRequest({ values })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0); // ReturnValue
        expect(response.readUInt16LE()).toBe(0); // PropertyProblemCount
        expect(response.hasMore()).toBe(false);

        expect(context.session.handles[5]?.draftProperties).toEqual({
            "55": "Hello",
            "3588": "to@example.com",
            "3587": "cc@example.com",
            "3586": "bcc@example.com",
            "4096": "Body text",
        });
    });

    it("Accepts but does not track an unsupported property, still reporting success.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1", draftProperties: {} };
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        const values: TaggedPropertyValue[] = [{ propertyId: 0x9999, propertyType: PropertyType.PtypInteger32, value: 42 }];
        handler.handle(new BufferReader(buildRequest({ values })), writer, context);

        const response = new BufferReader(writer.toBuffer());
        response.readUInt8();
        response.readUInt8();
        expect(response.readUInt32LE()).toBe(0);
        expect(context.session.handles[5]?.draftProperties).toEqual({});
    });

    it("Stringifies a tracked property's value via String() when it isn't already a PtypString/PtypString8.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftFolderUid: "folder:f1", draftProperties: {} };
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        const values: TaggedPropertyValue[] = [{ propertyId: 0x0037, propertyType: PropertyType.PtypInteger32, value: 42 }];
        handler.handle(new BufferReader(buildRequest({ values })), writer, context);

        expect(context.session.handles[5]?.draftProperties).toEqual({ "55": "42" });
    });

    it("Initializes draftProperties when setting properties on a handle that has none yet.", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "message:m1" };
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        const values: TaggedPropertyValue[] = [{ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hi" }];
        handler.handle(new BufferReader(buildRequest({ values })), writer, context);

        expect(context.session.handles[5]?.draftProperties).toEqual({ "55": "Hi" });
    });

    it("Throws if PropertyValueSize does not match the decoded PropertyValues length (malformed request).", () => {
        const context = makeContext();
        context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
        const handler = new RopSetPropertiesHandler();
        const writer = new BufferWriter();

        const malformed = buildRequest({ values: [{ propertyId: 0x0037, propertyType: PropertyType.PtypString, value: "Hi" }] });
        // Corrupt PropertyValueSize (bytes 2-3, little-endian) to no longer match the real encoded length.
        malformed.writeUInt16LE(9999, 2);

        expect(() => handler.handle(new BufferReader(malformed), writer, context)).toThrow(/PropertyValueSize/);
    });

    describe("Calendar named-property tracking", () => {
        it("Tracks PidTagMessageClass.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: 0x001a, propertyType: PropertyType.PtypString, value: "IPM.Appointment" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ "26": "IPM.Appointment" });
        });

        it("Tracks a recognized named Calendar property (PidLidLocation) when its ID was already assigned via RopGetPropertyIdsFromNames.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const locationId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_LOCATION });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: locationId, propertyType: PropertyType.PtypString, value: "Room 1" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ [String(locationId)]: "Room 1" });
        });

        it("Tracks a recognized PSETID_Common named Calendar property (PidLidReminderDelta).", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const deltaId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_COMMON, kind: "lid", lid: LID_REMINDER_DELTA });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: deltaId, propertyType: PropertyType.PtypInteger32, value: 15 }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ [String(deltaId)]: "15" });
        });

        it("Tracks PidLidGlobalObjectId (PSETID_Meeting) for meeting-response correlation.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const goidId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_MEETING, kind: "lid", lid: LID_GLOBAL_OBJECT_ID });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: goidId, propertyType: PropertyType.PtypBinary, value: Buffer.from([1, 2, 3]) }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ [String(goidId)]: Buffer.from([1, 2, 3]).toString("base64") });
        });

        it("Does not track an unrecognized LID under the PSETID_Meeting property set.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const id = assignOrGetNamedPropertyId(context.session, { guid: PSETID_MEETING, kind: "lid", lid: 0x9999 });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: id, propertyType: PropertyType.PtypString, value: "x" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({});
        });

        it("Does not track a named property whose ID was never assigned by RopGetPropertyIdsFromNames.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: 0x8000, propertyType: PropertyType.PtypString, value: "Untracked" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({});
        });

        it("Does not track a Kind=name named property (only Kind=LID Calendar properties are tracked).", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const id = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "name", name: "SomeCustomProp" });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: id, propertyType: PropertyType.PtypString, value: "x" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({});
        });

        it("Does not track a recognized LID under an unrecognized property-set GUID.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const id = assignOrGetNamedPropertyId(context.session, { guid: "11111111-0000-0000-c000-000000000046", kind: "lid", lid: LID_LOCATION });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: id, propertyType: PropertyType.PtypString, value: "x" }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({});
        });

        it("Does not track an unrecognized LID under a recognized PSETID_Common property set.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const id = assignOrGetNamedPropertyId(context.session, { guid: PSETID_COMMON, kind: "lid", lid: 0x9999 });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();

            const values: TaggedPropertyValue[] = [{ propertyId: id, propertyType: PropertyType.PtypInteger32, value: 1 }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({});
        });

        it("Stringifies a PtypTime value as an ISO-8601 string.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const startId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_START_WHOLE });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();
            const date = new Date("2026-09-07T14:00:00.000Z");

            const values: TaggedPropertyValue[] = [{ propertyId: startId, propertyType: PropertyType.PtypTime, value: date }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ [String(startId)]: date.toISOString() });
        });

        it("Stringifies a PtypBinary value as base64.", () => {
            const context = makeContext();
            context.session.handles[5] = { type: "message", entityUid: "", draftProperties: {} };
            const recurId = assignOrGetNamedPropertyId(context.session, { guid: PSETID_APPOINTMENT, kind: "lid", lid: LID_APPOINTMENT_RECUR });
            const handler = new RopSetPropertiesHandler();
            const writer = new BufferWriter();
            const blob = Buffer.from([1, 2, 3, 4]);

            const values: TaggedPropertyValue[] = [{ propertyId: recurId, propertyType: PropertyType.PtypBinary, value: blob }];
            handler.handle(new BufferReader(buildRequest({ values })), writer, context);

            expect(context.session.handles[5]?.draftProperties).toEqual({ [String(recurId)]: blob.toString("base64") });
        });
    });
});
