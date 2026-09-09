///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { assignOrGetFid, resolveFolderChildren, resolveFolderInfo } from "../../src/rop/FolderTarget.js";
import { MapiSessionContext } from "../../src/MapiSessionManager.js";

describe("FolderTarget Tests", () => {
    const MAILBOX_UID = "mailbox-1";

    describe("resolveFolderInfo", () => {
        it("Resolves a virtual folder to its display name, with zero counts and no children by default.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue([]), findOne: vi.fn() };
            const info = await resolveFolderInfo(MAILBOX_UID, "virtual:root", folderRepo as any);
            expect(info.displayName).toBe("Root");
            expect(info.unreadCount).toBe(0);
            expect(info.totalCount).toBe(0);
            expect(info.hasChildren).toBe(false);
        });

        it("Falls back to the raw name for a virtual folder with no known display-name mapping.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue([]), findOne: vi.fn() };
            const info = await resolveFolderInfo(MAILBOX_UID, "virtual:inbox", folderRepo as any);
            expect(info.displayName).toBe("inbox");
        });

        it("Resolves a real folder to its name/counts from the repo.", async () => {
            const folderRepo = {
                find: vi.fn().mockResolvedValue([]),
                findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox", unreadCount: 3, totalCount: 10 }),
            };
            const info = await resolveFolderInfo(MAILBOX_UID, "folder:f1", folderRepo as any);
            expect(info.displayName).toBe("Inbox");
            expect(info.unreadCount).toBe(3);
            expect(info.totalCount).toBe(10);
        });

        it("Degrades to empty values for a folder target whose real Folder has since vanished.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue([]), findOne: vi.fn().mockResolvedValue(undefined) };
            const info = await resolveFolderInfo(MAILBOX_UID, "folder:gone", folderRepo as any);
            expect(info.displayName).toBe("");
            expect(info.unreadCount).toBe(0);
            expect(info.totalCount).toBe(0);
        });

        it("Reports hasChildren true when resolveFolderChildren finds at least one.", async () => {
            const folderRepo = {
                find: vi.fn().mockResolvedValue([{ uid: "child1", parentFolderUid: "f1", mailboxUid: MAILBOX_UID }]),
                findOne: vi.fn().mockResolvedValue({ uid: "f1", name: "Inbox", unreadCount: 0, totalCount: 0 }),
            };
            const info = await resolveFolderInfo(MAILBOX_UID, "folder:f1", folderRepo as any);
            expect(info.hasChildren).toBe(true);
        });
    });

    describe("resolveFolderChildren", () => {
        const allFolders = [
            { uid: "top1", parentFolderUid: undefined, mailboxUid: MAILBOX_UID },
            { uid: "top2", parentFolderUid: undefined, mailboxUid: MAILBOX_UID },
            { uid: "child1", parentFolderUid: "top1", mailboxUid: MAILBOX_UID },
        ];

        it("Returns top-level real folders as children of virtual:root.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue(allFolders) };
            const children = await resolveFolderChildren(MAILBOX_UID, "virtual:root", folderRepo as any);
            expect(children.sort()).toEqual(["folder:top1", "folder:top2"]);
        });

        it("Treats virtual:ipmSubtree as an alias for the same top-level folder list as root.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue(allFolders) };
            const children = await resolveFolderChildren(MAILBOX_UID, "virtual:ipmSubtree", folderRepo as any);
            expect(children.sort()).toEqual(["folder:top1", "folder:top2"]);
        });

        it("Returns the real children of a real folder.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue(allFolders) };
            const children = await resolveFolderChildren(MAILBOX_UID, "folder:top1", folderRepo as any);
            expect(children).toEqual(["folder:child1"]);
        });

        it("Returns no children for a virtual folder other than root/ipmSubtree.", async () => {
            const folderRepo = { find: vi.fn().mockResolvedValue(allFolders) };
            const children = await resolveFolderChildren(MAILBOX_UID, "virtual:deferredAction", folderRepo as any);
            expect(children).toEqual([]);
        });
    });

    describe("assignOrGetFid", () => {
        it("Assigns sequential FIDs and reuses an existing assignment for the same target.", () => {
            const session = new MapiSessionContext({ mailboxUid: MAILBOX_UID, userUid: "user-1" });
            expect(assignOrGetFid(session, "virtual:root")).toBe(1);

            const newFid = assignOrGetFid(session, "folder:child1");
            expect(newFid).toBe(2);
            expect(session.folderIds["2"]).toBe("folder:child1");

            const reused = assignOrGetFid(session, "virtual:root");
            expect(reused).toBe(1);

            const reusedNew = assignOrGetFid(session, "folder:child1");
            expect(reusedNew).toBe(2);
        });

        it("Starts from FID 1 when the session has no prior assignments.", () => {
            const session = new MapiSessionContext({ mailboxUid: MAILBOX_UID, userUid: "user-1" });
            expect(assignOrGetFid(session, "folder:only")).toBe(1);
        });
    });
});
