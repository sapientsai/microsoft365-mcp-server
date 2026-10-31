import { describe, expect, it } from "vitest"

import { SET_TODO_TASK_STATE_PARAMS } from "../src/tools/todo-schemas"

describe("set_todo_task_state parameters", () => {
  it("accepts status, importance and an extension", () => {
    const parsed = SET_TODO_TASK_STATE_PARAMS.safeParse({
      list_id: "l",
      task_id: "t",
      status: "completed",
      importance: "high",
      extension: { name: "com.civala.assistant", fields: { rank: 1 } },
    })
    expect(parsed.success).toBe(true)
  })

  // zod would otherwise drop these silently, and the call would report success without them.
  it.each([
    ["a title", { title: "x" }],
    ["categories", { categories: ["a"] }],
    ["a due date", { due_date: "2026-11-02" }],
  ])("refuses %s rather than ignoring it", (_label, extra) => {
    const parsed = SET_TODO_TASK_STATE_PARAMS.safeParse({ list_id: "l", task_id: "t", status: "completed", ...extra })
    expect(parsed.success).toBe(false)
  })

  it("refuses a status or importance Graph doesn't have", () => {
    expect(SET_TODO_TASK_STATE_PARAMS.safeParse({ list_id: "l", task_id: "t", status: "done" }).success).toBe(false)
    expect(SET_TODO_TASK_STATE_PARAMS.safeParse({ list_id: "l", task_id: "t", importance: "urgent" }).success).toBe(
      false,
    )
  })
})
