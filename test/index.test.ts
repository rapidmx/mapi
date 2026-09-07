///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Smoke test for the `mapi` barrel module (`@rapidrest/mail/mapi`) - every other MAPI test imports its subject
// directly from its concrete file, so this file's only job is to catch a typo'd or missing `export *` in
// src/mapi/index.ts itself: one representative, real symbol from each re-exported file, confirmed present and
// of the expected shape. See test/eas/index.test.ts for the identical precedent this follows.
import * as Mapi from "../src/index.js";

describe("mapi/index Tests", () => {
    it("Re-exports the wire-format codecs.", () => {
        expect(typeof Mapi.BufferReader).toBe("function");
        expect(typeof Mapi.BufferWriter).toBe("function");
        expect(typeof Mapi.encodeGuid).toBe("function");
        expect(typeof Mapi.decodeGuid).toBe("function");
        expect(typeof Mapi.readPropertyTag).toBe("function");
        expect(typeof Mapi.writePropertyTag).toBe("function");
        expect(typeof Mapi.decodeRopBuffer).toBe("function");
        expect(typeof Mapi.encodeRopBuffer).toBe("function");
        expect(typeof Mapi.encodeAppointmentRecurrence).toBe("function");
        expect(typeof Mapi.decodeAppointmentRecurrence).toBe("function");
        expect(typeof Mapi.encodeTimeZoneStruct).toBe("function");
        expect(typeof Mapi.decodeTimeZoneStruct).toBe("function");
        expect(typeof Mapi.encodeGlobalObjectId).toBe("function");
        expect(typeof Mapi.decodeGlobalObjectId).toBe("function");
    });

    it("Re-exports the transport base classes, session manager, and ROP dispatcher.", () => {
        expect(typeof Mapi.BaseMapiEmsmdbRoute).toBe("function");
        expect(typeof Mapi.BaseMapiNspiRoute).toBe("function");
        expect(typeof Mapi.MapiSessionManager).toBe("function");
        expect(typeof Mapi.dispatchRops).toBe("function");
    });

    it("Re-exports every ROP handler class.", () => {
        expect(typeof Mapi.RopLogonHandler).toBe("function");
        expect(typeof Mapi.RopReleaseHandler).toBe("function");
        expect(typeof Mapi.RopOpenFolderHandler).toBe("function");
        expect(typeof Mapi.RopGetHierarchyTableHandler).toBe("function");
        expect(typeof Mapi.RopSetColumnsHandler).toBe("function");
        expect(typeof Mapi.RopQueryRowsHandler).toBe("function");
        expect(typeof Mapi.RopGetContentsTableHandler).toBe("function");
        expect(typeof Mapi.RopOpenMessageHandler).toBe("function");
        expect(typeof Mapi.RopGetPropertiesSpecificHandler).toBe("function");
        expect(typeof Mapi.RopOpenStreamHandler).toBe("function");
        expect(typeof Mapi.RopReadStreamHandler).toBe("function");
        expect(typeof Mapi.RopCreateMessageHandler).toBe("function");
        expect(typeof Mapi.RopSetPropertiesHandler).toBe("function");
        expect(typeof Mapi.RopWriteStreamHandler).toBe("function");
        expect(typeof Mapi.RopSaveChangesMessageHandler).toBe("function");
        expect(typeof Mapi.RopSubmitMessageHandler).toBe("function");
        expect(typeof Mapi.RopGetPropertyIdsFromNamesHandler).toBe("function");
        expect(typeof Mapi.RopDeleteMessagesHandler).toBe("function");
        expect(typeof Mapi.RopDeleteFolderHandler).toBe("function");
        expect(typeof Mapi.RopFastTransferSourceCopyToHandler).toBe("function");
        expect(typeof Mapi.RopFastTransferSourceCopyPropertiesHandler).toBe("function");
        expect(typeof Mapi.RopFastTransferSourceGetBufferHandler).toBe("function");
    });

    it("Re-exports the NSPI codec and handlers.", () => {
        expect(typeof Mapi.readStat).toBe("function");
        expect(typeof Mapi.writeStat).toBe("function");
        expect(typeof Mapi.handleNspiBind).toBe("function");
        expect(typeof Mapi.handleNspiUnbind).toBe("function");
        expect(typeof Mapi.handleNspiGetMatches).toBe("function");
    });
});
