///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { assignOrGetMid, resolveFolderMessages, resolveMessageInfo } from "../../src/rop/MessageTarget.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

describe("MessageTarget Tests", () => {
    describe("resolveMessageInfo", () => {
        it("Resolves a real message's subject/read/hasAttachments/receivedDate/labelUids.", async () => {
            const receivedDate = new Date("2026-01-01T00:00:00.000Z");
            const messageRepo = {
                findOne: vi.fn().mockResolvedValue({
                    uid: "m1",
                    subject: "Hello",
                    flags: { read: true },
                    hasAttachments: true,
                    receivedDate,
                    labelUids: ["label1"],
                }),
            };
            const info = await resolveMessageInfo("message:m1", messageRepo as any);
            expect(info).toEqual({ subject: "Hello", read: true, hasAttachments: true, receivedDate, labelUids: ["label1"] });
            expect(messageRepo.findOne).toHaveBeenCalledWith("m1", { ignoreACL: true });
        });

        it("Degrades to empty values for a message target whose real Message has since vanished.", async () => {
            const messageRepo = { findOne: vi.fn().mockResolvedValue(undefined) };
            const info = await resolveMessageInfo("message:gone", messageRepo as any);
            expect(info.subject).toBe("");
            expect(info.read).toBe(false);
            expect(info.hasAttachments).toBe(false);
            expect(info.receivedDate).toEqual(new Date(0));
            expect(info.labelUids).toEqual([]);
        });

        it("Defaults labelUids to an empty array when the message has none set.", async () => {
            const messageRepo = { findOne: vi.fn().mockResolvedValue({ uid: "m1", subject: "Hi" }) };
            const info = await resolveMessageInfo("message:m1", messageRepo as any);
            expect(info.labelUids).toEqual([]);
        });
    });

    describe("resolveFolderMessages", () => {
        it("Returns each message as a message:<uid> target string.", async () => {
            const messageRepo = { find: vi.fn().mockResolvedValue([{ uid: "m1" }, { uid: "m2" }]) };
            const targets = await resolveFolderMessages("folder-1", messageRepo as any);
            expect(targets).toEqual(["message:m1", "message:m2"]);
            expect(messageRepo.find).toHaveBeenCalledWith({ folderUid: "folder-1" }, { ignoreACL: true });
        });
    });

    describe("assignOrGetMid", () => {
        it("Assigns sequential MIDs and reuses an existing assignment for the same target.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            expect(assignOrGetMid(session, "message:m1")).toBe(1);

            const newMid = assignOrGetMid(session, "message:m2");
            expect(newMid).toBe(2);
            expect(session.messageIds["2"]).toBe("message:m2");

            const reused = assignOrGetMid(session, "message:m1");
            expect(reused).toBe(1);
        });

        it("Starts from MID 1 when the session has no prior assignments.", () => {
            const session = new MapiSessionContext({ mailboxUid: "mailbox-1", userUid: "user-1" });
            expect(assignOrGetMid(session, "message:only")).toBe(1);
        });
    });
});
