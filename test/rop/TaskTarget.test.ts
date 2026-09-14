///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { resolveTaskInfo } from "../../src/rop/TaskTarget.js";

describe("TaskTarget Tests", () => {
    describe("resolveTaskInfo", () => {
        it("Resolves a real task's fields.", async () => {
            const dueDate = new Date("2026-09-10T00:00:00.000Z");
            const taskRepo = { findOne: vi.fn().mockResolvedValue({ uid: "t1", title: "Ship it", completed: false, dueDate }) };

            const info = await resolveTaskInfo("task:t1", taskRepo as any);

            expect(info).toEqual({ title: "Ship it", completed: false, dueDate });
            expect(taskRepo.findOne).toHaveBeenCalledWith("t1", { ignoreACL: true });
        });

        it("Degrades to empty-looking values for a task target whose real Task has since vanished.", async () => {
            const taskRepo = { findOne: vi.fn().mockResolvedValue(undefined) };

            const info = await resolveTaskInfo("task:gone", taskRepo as any);

            expect(info.title).toBe("");
            expect(info.completed).toBe(false);
            expect(info.dueDate).toBeUndefined();
        });
    });
});
