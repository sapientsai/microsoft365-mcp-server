import { Some } from "functype"
import { Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createTodoTask, updateTodoTask } from "../src/tools/todo-tools"

const mockClient = {
  createTodoTask: vi.fn(),
  updateTodoTask: vi.fn(),
}

const LIST_ID = "list-1"

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("todo-tools", () => {
  describe("createTodoTask", () => {
    it("should create a task with just a title", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      const result = await createTodoTask({ list_id: LIST_ID, title: "Buy milk" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, { title: "Buy milk" })
    })

    // Graph's todoTask.body is an itemBody whose bodyType enum is lowercase ("text" |
    // "html"). Sending "Text" is rejected with "Requested value 'Text' was not found",
    // which fails the whole create — not just the body.
    it("should send a lowercase bodyType so Graph accepts the body", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Call plumber" }))
      await createTodoTask({ list_id: LIST_ID, title: "Call plumber", body: "0400 000 000" })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Call plumber",
        body: { contentType: "text", content: "0400 000 000" },
      })
    })

    it("should pass due date and importance through", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Renew rego" }))
      await createTodoTask({
        list_id: LIST_ID,
        title: "Renew rego",
        due_date: "2026-10-01T00:00:00",
        importance: "high",
      })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Renew rego",
        dueDateTime: { dateTime: "2026-10-01T00:00:00", timeZone: "UTC" },
        importance: "high",
      })
    })

    it("should surface a create failure as a UserError", async () => {
      mockClient.createTodoTask.mockResolvedValue(
        (await import("functype/either")).Left({ message: "Requested value 'Text' was not found." }),
      )
      const result = await createTodoTask({ list_id: LIST_ID, title: "Nope", body: "x" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to create task")
    })
  })

  describe("updateTodoTask", () => {
    it("should send a lowercase bodyType on update too", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", title: "Call plumber" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", body: "new notes" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        body: { contentType: "text", content: "new notes" },
      })
    })

    it("should only send the fields that were provided", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", status: "completed" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", { status: "completed" })
    })
  })
})
