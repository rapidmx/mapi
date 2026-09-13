///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../../src/codec/BufferCursor.js";
import { PropertyType, readPropertyValue, writePropertyTag, writeTaggedPropertyValue } from "../../src/codec/PropertyValue.js";
import { handleNspiGetMatches } from "../../src/nspi/NspiGetMatchesHandler.js";
import { writeStat, type Stat } from "../../src/nspi/NspiCodec.js";

function makeRes() {
    return { status: vi.fn().mockReturnThis(), send: vi.fn().mockReturnThis() };
}

const SAMPLE_STAT: Stat = {
    sortType: 0,
    containerId: 0,
    currentRec: 0,
    delta: 0,
    numPos: 0,
    totalRecs: 0,
    codePage: 0,
    templateLocale: 0,
    sortLocale: 0,
};

function buildRequest(build: (writer: BufferWriter) => void): { headers: {}; rawBody: Buffer } {
    const writer = new BufferWriter();
    build(writer);
    return { headers: {}, rawBody: writer.toBuffer() };
}

describe("handleNspiGetMatches Tests", () => {
    it("Decodes HasState/State and HasMinimalIds/MinimalIds fields to advance the reader correctly, without honoring them.", async () => {
        const contactRepo = { find: vi.fn().mockResolvedValue([]) };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0); // Reserved
            writer.writeUInt8(1); // HasState
            writeStat(writer, SAMPLE_STAT);
            writer.writeUInt8(1); // HasMinimalIds
            writer.writeUInt32LE(2); // MinimalIdCount
            writer.writeUInt32LE(111);
            writer.writeUInt32LE(222);
            writer.writeUInt32LE(0); // InterfaceOptionFlags
            writer.writeUInt8(0); // HasFilter
            writer.writeUInt8(0); // HasPropertyName
            writer.writeUInt32LE(50); // RowCount
            writer.writeUInt8(0); // HasColumns
            writer.writeUInt32LE(0); // AuxiliaryBufferSize
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        expect(contactRepo.find).toHaveBeenCalledWith({ mailboxUid: "mailbox-1" }, { ignoreACL: true });
        expect(res.status).toHaveBeenCalledWith(200);
    });

    it("Decodes HasPropertyName to advance the reader correctly, without honoring it.", async () => {
        const contactRepo = { find: vi.fn().mockResolvedValue([]) };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0); // HasState
            writer.writeUInt8(0); // HasMinimalIds
            writer.writeUInt32LE(0);
            writer.writeUInt8(0); // HasFilter
            writer.writeUInt8(1); // HasPropertyName
            writer.writeBytes(Buffer.alloc(16)); // PropertyNameGuid
            writer.writeUInt32LE(1); // PropertyNameId
            writer.writeUInt32LE(50); // RowCount
            writer.writeUInt8(0); // HasColumns
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        expect(res.status).toHaveBeenCalledWith(200);
    });

    it("Uses an explicit include column list from HasColumns instead of the default column set.", async () => {
        const contactRepo = { find: vi.fn().mockResolvedValue([{ uid: "c1", displayName: "Jane", emails: [] }]) };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
            writer.writeUInt8(0); // HasFilter
            writer.writeUInt8(0); // HasPropertyName
            writer.writeUInt32LE(50); // RowCount
            writer.writeUInt8(1); // HasColumns
            writer.writeUInt32LE(1); // PropertyTagCount
            writePropertyTag(writer, { propertyId: 0x0e07, propertyType: PropertyType.PtypInteger32 }); // MessageFlags - not DisplayName/EmailAddress
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        const body = res.send.mock.calls[0][0] as Buffer;
        const reader = new BufferReader(body);
        reader.readUInt32LE(); // StatusCode
        reader.readUInt32LE(); // ErrorCode
        reader.readUInt8(); // HasState
        reader.readBytes(36); // State
        reader.readUInt8(); // HasMinimalIds
        reader.readUInt32LE(); // MinimalIdCount
        reader.readUInt32LE(); // the one MinimalId
        reader.readUInt8(); // HasColumnsAndRows
        const columnCount = reader.readUInt32LE();
        expect(columnCount).toBe(1);
        const propertyType = reader.readUInt16LE();
        const propertyId = reader.readUInt16LE();
        expect(propertyId).toBe(0x0e07);
        reader.readUInt32LE(); // RowCount
        reader.readUInt8(); // AddressBookPropertyRow's own Flags
        // MessageFlags is a fixed-size PtypInteger32 - no HasValue byte, falls back to defaultValueForType's 0
        // since this column isn't DisplayName/EmailAddress.
        expect(readPropertyValue(reader, propertyType)).toBe(0);
    });

    it("Does not throw when a contact's own emails field is missing entirely (not just empty), degrading EmailAddress to empty string.", async () => {
        const contactRepo = { find: vi.fn().mockResolvedValue([{ uid: "c1", displayName: "No Emails Field" }]) };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
            writer.writeUInt8(0); // HasFilter
            writer.writeUInt8(0); // HasPropertyName
            writer.writeUInt32LE(50);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await expect(handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any)).resolves.not.toThrow();
        expect(res.status).toHaveBeenCalledWith(200);
    });

    it("Reports HasMinimalIds=0 and HasColumnsAndRows=0 (no MinimalIdCount/RowCount/RowData at all) when there are zero matches.", async () => {
        const contactRepo = { find: vi.fn().mockResolvedValue([]) };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
            writer.writeUInt8(0); // HasFilter
            writer.writeUInt8(0); // HasPropertyName
            writer.writeUInt32LE(50);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        const body = res.send.mock.calls[0][0] as Buffer;
        const reader = new BufferReader(body);
        reader.readUInt32LE();
        reader.readUInt32LE();
        reader.readUInt8();
        reader.readBytes(36);
        expect(reader.readUInt8()).toBe(0); // HasMinimalIds
        expect(reader.readUInt8()).toBe(0); // HasColumnsAndRows
        expect(reader.readUInt32LE()).toBe(0); // AuxiliaryBufferSize
        expect(reader.hasMore()).toBe(false);
    });

    it("Throws (via BufferReader) when req.rawBody is undefined, falling back to an empty buffer.", async () => {
        const contactRepo = { find: vi.fn() };
        const req = { headers: {}, rawBody: undefined } as any;
        await expect(handleNspiGetMatches(req, makeRes() as any, "mailbox-1", contactRepo as any)).rejects.toThrow();
    });

    it("Filters via a ContentRestriction search term, querying displayName/givenName/surname/company, merging by uid, and sorting by displayName.", async () => {
        const find = vi.fn().mockImplementation((query: any) => {
            if ("displayName" in query) return Promise.resolve([{ uid: "c1", displayName: "Zed", emails: [] }]);
            if ("surname" in query) return Promise.resolve([{ uid: "c2", displayName: "Ann", emails: [] }]);
            return Promise.resolve([]);
        });
        const contactRepo = { find };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
            writer.writeUInt8(1); // HasFilter
            writer.writeUInt8(0x03); // RestrictType - ContentRestriction
            writer.writeUInt16LE(0x0001);
            writer.writeUInt16LE(0x0001);
            writePropertyTag(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString });
            writeTaggedPropertyValue(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString, value: "jane" });
            writer.writeUInt8(0); // HasPropertyName
            writer.writeUInt32LE(50);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        expect(find).toHaveBeenCalledWith({ mailboxUid: "mailbox-1", displayName: "regex(jane)" }, { ignoreACL: true });
        expect(find).toHaveBeenCalledTimes(4); // displayName/givenName/surname/company
        const body = res.send.mock.calls[0][0] as Buffer;
        const reader = new BufferReader(body);
        reader.readUInt32LE();
        reader.readUInt32LE();
        reader.readUInt8();
        reader.readBytes(36);
        expect(reader.readUInt8()).toBe(1); // HasMinimalIds
        expect(reader.readUInt32LE()).toBe(2); // MinimalIdCount - Zed (displayName match) + Ann (surname match), merged
        reader.readUInt32LE(); // Ann's MinimalId
        reader.readUInt32LE(); // Zed's MinimalId
        reader.readUInt8(); // HasColumnsAndRows
        const columnCount = reader.readUInt32LE();
        for (let i = 0; i < columnCount; i++) {
            reader.readUInt16LE();
            reader.readUInt16LE();
        }
        reader.readUInt32LE(); // RowCount
        // Sorted by displayName ("Ann" before "Zed") - the first row's DisplayName column proves the comparator ran.
        reader.readUInt8(); // AddressBookPropertyRow's own Flags
        reader.readUInt8(); // HasValue
        expect(reader.readNullTerminatedUtf16LE()).toBe("Ann");
    });

    it("Escapes regex metacharacters in the search term before building the regex() query pattern, so a literal '.' matches only itself.", async () => {
        const find = vi.fn().mockResolvedValue([]);
        const contactRepo = { find };
        const req = buildRequest((writer) => {
            writer.writeUInt32LE(0);
            writer.writeUInt8(0);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
            writer.writeUInt8(1); // HasFilter
            writer.writeUInt8(0x03);
            writer.writeUInt16LE(0x0001);
            writer.writeUInt16LE(0x0001);
            writePropertyTag(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString });
            writeTaggedPropertyValue(writer, { propertyId: 0x3001, propertyType: PropertyType.PtypString, value: "a.b" });
            writer.writeUInt8(0);
            writer.writeUInt32LE(50);
            writer.writeUInt8(0);
            writer.writeUInt32LE(0);
        });
        const res = makeRes();

        await handleNspiGetMatches(req as any, res as any, "mailbox-1", contactRepo as any);

        expect(find).toHaveBeenCalledWith({ mailboxUid: "mailbox-1", displayName: "regex(a\\.b)" }, { ignoreACL: true });
    });
});
