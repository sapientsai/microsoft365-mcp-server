import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createTodoList, createTodoTask, deleteTodoTask, listTodoTasks, updateTodoTask } from "../src/tools/todo-tools"

const mockClient = {
  createTodoTask: vi.fn(),
  updateTodoTask: vi.fn(),
  getTodoTask: vi.fn(),
  deleteTodoTask: vi.fn(),
  listTodoLists: vi.fn(),
  createTodoList: vi.fn(),
  requestPaginated: vi.fn(),
  listTodoTasks: vi.fn(),
  listAllTodoTasks: vi.fn(),
}

const QUARTERLY = {
  pattern: { type: "absoluteMonthly", interval: 3, dayOfMonth: 1 },
  range: { type: "noEnd", startDate: "2026-10-01" },
}

const LIST_ID = "list-1"

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("todo-tools", () => {
  describe("createTodoList", () => {
    it("should create a list when the name is free", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "l1", displayName: "Household" }]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      const result = await createTodoList({ display_name: "Renovation" })

      expect(result.isRight()).toBe(true)
      expect(mockClient.createTodoList).toHaveBeenCalledWith("Renovation")
    })

    // Graph happily creates a second list with the same displayName, and no client can
    // then tell the two apart — so tasks scatter across both.
    it("should refuse a name an existing list already uses, case-insensitively", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "l1", displayName: "Household" }]))

      const result = await createTodoList({ display_name: "household" })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should check every page of lists, not just the first", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      await createTodoList({ display_name: "Renovation" })

      expect(mockClient.requestPaginated).toHaveBeenCalledWith("/me/todo/lists")
      expect(mockClient.listTodoLists).not.toHaveBeenCalled()
    })

    // A read failure leaves the duplicate check blind, so creating anyway could still
    // produce the duplicate the guard exists to prevent.
    it("should refuse to create when the existing lists cannot be read", async () => {
      mockClient.requestPaginated.mockResolvedValue(Left({ message: "Service unavailable" }))

      const result = await createTodoList({ display_name: "Renovation" })

      expect(result.isLeft()).toBe(true)
      expect(
        result.fold(
          (error) => error.message,
          () => "",
        ),
      ).toContain("Service unavailable")
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should refuse a blank name without calling Graph", async () => {
      const result = await createTodoList({ display_name: "   " })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.requestPaginated).not.toHaveBeenCalled()
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should trim the name before sending it", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      await createTodoList({ display_name: "  Renovation  " })

      expect(mockClient.createTodoList).toHaveBeenCalledWith("Renovation")
    })
  })

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

    it("should attach a recurrence and default its range to the due date", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters" }))
      await createTodoTask({
        list_id: LIST_ID,
        title: "Clear the gutters",
        due_date: "2026-10-01T00:00:00",
        recurrence: { pattern: "absoluteMonthly", interval: 3, day_of_month: 1 },
      })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Clear the gutters",
        dueDateTime: { dateTime: "2026-10-01T00:00:00", timeZone: "UTC" },
        recurrence: {
          pattern: { type: "absoluteMonthly", interval: 3, dayOfMonth: 1 },
          range: { type: "noEnd", startDate: "2026-10-01" },
        },
      })
    })

    // To Do rolls a recurring task forward from its due date, so without one the
    // task repeats but never surfaces in Today.
    it("should refuse a recurring task with no due date", async () => {
      const result = await createTodoTask({
        list_id: LIST_ID,
        title: "Clear the gutters",
        recurrence: { pattern: "daily" },
      })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("due_date")
      expect(mockClient.createTodoTask).not.toHaveBeenCalled()
    })

    it("should not call Graph when the recurrence is incomplete", async () => {
      const result = await createTodoTask({
        list_id: LIST_ID,
        title: "Bin night",
        due_date: "2026-10-01T00:00:00",
        recurrence: { pattern: "weekly" },
      })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("days_of_week")
      expect(mockClient.createTodoTask).not.toHaveBeenCalled()
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

    it("should add a recurrence to an existing task", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1" }))
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        recurrence: { pattern: "absoluteYearly", day_of_month: 1, month: 10, start_date: "2026-10-01" },
      })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        recurrence: {
          pattern: { type: "absoluteYearly", interval: 1, dayOfMonth: 1, month: 10 },
          range: { type: "noEnd", startDate: "2026-10-01" },
        },
      })
    })

    // Graph only clears a recurrence on an explicit null; omitting it leaves the
    // existing pattern untouched.
    it("should clear a recurrence with an explicit null", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", clear_recurrence: true })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", { recurrence: null })
    })

    // Graph rolls a recurring task forward on completion: the same id comes back as
    // notStarted with the next due date, and the completed occurrence becomes a
    // separate task. "Task updated" over a notStarted task reads as a failure.
    it("should explain the roll-forward when completing a recurring task", async () => {
      mockClient.updateTodoTask.mockResolvedValue(
        Right({
          id: "t1",
          title: "Clear the gutters",
          status: "notStarted",
          dueDateTime: { dateTime: "2027-01-01T00:00:00.0000000" },
          recurrence: QUARTERLY,
        }),
      )
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(result.isRight()).toBe(true)
      const text = result.value as string
      expect(text).toContain("Occurrence completed")
      expect(text).toContain("2027-01-01")
    })

    it("should say plainly when a non-recurring task is completed", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk", status: "completed" }))
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(result.value as string).toContain("Task completed.")
    })

    it("should reject setting and clearing a recurrence at once", async () => {
      const result = await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        clear_recurrence: true,
        recurrence: { pattern: "daily", start_date: "2026-10-01" },
      })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })
  })

  // A fake To Do list that behaves as Graph did live on 2026-10-28, for daily tasks:
  // a dueDateTime PATCH on a repeating task — with or without a recurrence beside
  // it — rolls the series forward a day and creates a second task on the requested
  // date; a recurrence PATCH moves the due date to the range's startDate.
  const fakeTodoList = (initial: ReadonlyArray<Record<string, unknown>>) => {
    const tasks = new Map(initial.map((t) => [t.id as string, { ...t }]))
    const dueOf = (date: string) => ({ dateTime: `${date}T00:00:00.0000000`, timeZone: "UTC" })
    const nextDay = (date: string) => new Date(Date.parse(`${date}T00:00:00Z`) + 864e5).toISOString().slice(0, 10)
    mockClient.getTodoTask.mockImplementation(async (_list: string, id: string) => Right({ ...tasks.get(id) }))
    mockClient.updateTodoTask.mockImplementation(async (_list: string, id: string, patch: Record<string, unknown>) => {
      const task = tasks.get(id)!
      const due = patch.dueDateTime as { dateTime: string } | undefined
      if (due && task.recurrence && patch.recurrence !== null) {
        tasks.set(`${id}-copy`, { ...task, id: `${id}-copy`, dueDateTime: dueOf(due.dateTime.slice(0, 10)) })
        const current = (task.dueDateTime as { dateTime: string }).dateTime.slice(0, 10)
        Object.assign(task, { dueDateTime: dueOf(nextDay(current)) })
        return Right({ ...task })
      }
      Object.assign(task, patch)
      if (due) task.dueDateTime = dueOf(due.dateTime.slice(0, 10))
      const recurrence = patch.recurrence as { range: { startDate: string } } | null | undefined
      if (recurrence) task.dueDateTime = dueOf(recurrence.range.startDate)
      return Right({ ...task })
    })
    return tasks
  }

  const DAILY = { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-10-30" } }
  const dueDates = (tasks: Map<string, Record<string, unknown>>) =>
    [...tasks.values()].map((t) => (t.dueDateTime as { dateTime: string }).dateTime.slice(0, 10))

  describe("updateTodoTask rescheduling a repeating task", () => {
    const repeating = () => [
      { id: "t1", title: "Water plants", dueDateTime: { dateTime: "2026-10-30T00:00:00.0000000" }, recurrence: DAILY },
    ]

    it("moves it with due_date alone, without creating a second task", async () => {
      const tasks = fakeTodoList(repeating())
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06T12:00:00Z" })
      expect(result.isRight()).toBe(true)
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.updateTodoTask.mock.calls[0]![2]).not.toHaveProperty("dueDateTime")
    })

    it("moves it with due_date and recurrence together, without creating a second task", async () => {
      const tasks = fakeTodoList(repeating())
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        recurrence: { pattern: "daily" },
      })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.getTodoTask).not.toHaveBeenCalled()
    })

    it("keeps the rest of the existing pattern when re-anchoring it", async () => {
      fakeTodoList([{ ...repeating()[0], recurrence: QUARTERLY }])
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2027-01-01" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        recurrence: { pattern: QUARTERLY.pattern, range: { type: "noEnd", startDate: "2027-01-01" } },
      })
    })

    it("still sends dueDateTime when the recurrence is being cleared", async () => {
      const tasks = fakeTodoList(repeating())
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        clear_recurrence: true,
      })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.getTodoTask).not.toHaveBeenCalled()
    })

    it("sends dueDateTime for a one-off task", async () => {
      const tasks = fakeTodoList([
        { id: "t1", title: "Buy milk", dueDateTime: { dateTime: "2026-10-30T00:00:00.0000000" } },
      ])
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06T12:00:00Z" })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.updateTodoTask.mock.calls[0]![2]).toHaveProperty("dueDateTime")
    })

    // Graph puts the task on the first date the pattern allows from startDate.
    it("says where the task landed when the pattern moved it off the asked date", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", recurrence: QUARTERLY }))
      mockClient.updateTodoTask.mockResolvedValue(
        Right({ id: "t1", dueDateTime: { dateTime: "2027-01-01T00:00:00.0000000" }, recurrence: QUARTERLY }),
      )
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-12-15" })
      const text = result.value as string
      expect(text).toContain("Graph put the task on 2027-01-01, not 2026-12-15.")
      expect(text).not.toContain("so the next occurrence is")
    })

    it("refuses a due_date that differs from recurrence.start_date, without calling Graph", async () => {
      const result = await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        recurrence: { pattern: "daily", start_date: "2026-11-09" },
      })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("2026-11-06")
      expect((result.value as Error).message).toContain("2026-11-09")
      expect(mockClient.getTodoTask).not.toHaveBeenCalled()
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })

    it("accepts a due_date and recurrence.start_date naming the same day", async () => {
      const tasks = fakeTodoList(repeating())
      const result = await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        recurrence: { pattern: "daily", start_date: "2026-11-06" },
      })
      expect(result.isRight()).toBe(true)
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
    })

    it("refuses to move an endDate series past its end, without updating", async () => {
      fakeTodoList([
        {
          ...repeating()[0],
          recurrence: { ...DAILY, range: { type: "endDate", startDate: "2026-10-30", endDate: "2026-11-03" } },
        },
      ])
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("stops repeating on 2026-11-03")
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })

    it("moves an endDate series within its end, keeping the end", async () => {
      const range = { type: "endDate", startDate: "2026-10-30", endDate: "2026-12-31" }
      fakeTodoList([{ ...repeating()[0], recurrence: { ...DAILY, range } }])
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        recurrence: { pattern: DAILY.pattern, range: { ...range, startDate: "2026-11-06" } },
      })
    })

    // Graph echoes endDate 0001-01-01 on a noEnd range; that is not an end date.
    it("moves a noEnd series whose echoed endDate is 0001-01-01", async () => {
      const range = { type: "noEnd", startDate: "2026-10-30", endDate: "0001-01-01", numberOfOccurrences: 0 }
      const tasks = fakeTodoList([{ ...repeating()[0], recurrence: { ...DAILY, range } }])
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06" })
      expect(result.isRight()).toBe(true)
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
    })

    it("gives one next-occurrence sentence when completing and moving at once", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", recurrence: QUARTERLY }))
      mockClient.updateTodoTask.mockResolvedValue(
        Right({
          id: "t1",
          status: "notStarted",
          dueDateTime: { dateTime: "2027-04-01T00:00:00.0000000" },
          recurrence: QUARTERLY,
        }),
      )
      const result = await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        status: "completed",
        due_date: "2026-12-15",
      })
      const text = result.value as string
      expect(text).toContain("Occurrence completed")
      expect(text).not.toContain("first date")
    })

    it("does not update when the task cannot be read first", async () => {
      mockClient.getTodoTask.mockResolvedValue(Left({ message: "ErrorItemNotFound" }))
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })
  })

  describe("deleteTodoTask", () => {
    it("should delete a one-off task and name it", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      mockClient.deleteTodoTask.mockResolvedValue(Right({}))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isRight()).toBe(true)
      expect(result.value as string).toContain('Deleted "Buy milk"')
      expect(mockClient.deleteTodoTask).toHaveBeenCalledWith(LIST_ID, "t1")
    })

    // A task id addresses the whole series, not one occurrence, and Graph has no
    // restore — so this refusal is the only thing standing between a routine
    // tidy-up and silently destroying a recurring chore.
    it("should refuse a recurring task without force", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters", recurrence: QUARTERLY }))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isLeft()).toBe(true)
      const message = (result.value as Error).message
      expect(message).toContain("ends the whole series")
      expect(message).toContain("every 3 months")
      expect(message).toContain("clear_recurrence")
      expect(mockClient.deleteTodoTask).not.toHaveBeenCalled()
    })

    it("should delete a recurring task when forced, and say the series ended", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters", recurrence: QUARTERLY }))
      mockClient.deleteTodoTask.mockResolvedValue(Right({}))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1", force: true })
      expect(result.isRight()).toBe(true)
      expect(result.value as string).toContain("whole recurring series")
      expect(mockClient.deleteTodoTask).toHaveBeenCalledWith(LIST_ID, "t1")
    })

    it("should not delete when the task cannot be read first", async () => {
      mockClient.getTodoTask.mockResolvedValue((await import("functype/either")).Left({ message: "ErrorItemNotFound" }))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "gone" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("before deleting")
      expect(mockClient.deleteTodoTask).not.toHaveBeenCalled()
    })

    it("should surface a delete failure as a UserError", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      mockClient.deleteTodoTask.mockResolvedValue(
        (await import("functype/either")).Left({ message: "ErrorAccessDenied" }),
      )
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to delete task")
    })
  })
})

describe("list_todo_tasks output", () => {
  const task = {
    id: "t1",
    title: "Reply to Bel",
    status: "notStarted",
    importance: "high",
    dueDateTime: { dateTime: "2026-11-02T00:00:00.0000000", timeZone: "UTC" },
    lastModifiedDateTime: "2026-10-30T12:00:00Z",
    categories: ["Red category"],
    linkedResources: [{ externalId: "civala:mail:AAMk1", applicationName: "Outlook" }],
  }

  it("expands linkedResources and prints every field in order", async () => {
    mockClient.listTodoTasks.mockResolvedValue(Right({ value: [task] }))
    const result = await listTodoTasks({ list_id: LIST_ID })
    expect(mockClient.listTodoTasks).toHaveBeenCalledWith(LIST_ID, { $expand: ["linkedResources"] })
    expect(result.value).toContain(
      "- **Reply to Bel** [notStarted] [high] (Due: 2026-11-02T00:00:00.0000000) (Modified: 2026-10-30T12:00:00Z) " +
        "(Categories: Red category) (Link: civala:mail:AAMk1) (ID: t1)",
    )
  })

  it("leaves out what Graph didn't return instead of filling in defaults", async () => {
    mockClient.listTodoTasks.mockResolvedValue(Right({ value: [{ id: "t2" }] }))
    const result = await listTodoTasks({ list_id: LIST_ID })
    expect(result.value).toContain("- **** (ID: t2)")
    expect(result.value).not.toContain("Untitled")
    expect(result.value).not.toContain("notStarted")
  })

  it("shows importance only when it isn't normal", async () => {
    mockClient.listTodoTasks.mockResolvedValue(Right({ value: [{ ...task, importance: "normal" }] }))
    const result = await listTodoTasks({ list_id: LIST_ID })
    expect(result.value).toContain("- **Reply to Bel** [notStarted] (Due:")
  })

  it("prints the completion time with its time zone", async () => {
    mockClient.listTodoTasks.mockResolvedValue(
      Right({ value: [{ id: "t3", completedDateTime: { dateTime: "2026-10-29T00:00:00.0000000", timeZone: "UTC" } }] }),
    )
    const result = await listTodoTasks({ list_id: LIST_ID })
    expect(result.value).toContain("(Completed: 2026-10-29T00:00:00.0000000 UTC) (ID: t3)")
  })

  // A flagged email's task title is the email's subject, chosen by its sender.
  it("keeps a title on one line, and a category or link from ending its group early", async () => {
    mockClient.listTodoTasks.mockResolvedValue(
      Right({
        value: [
          {
            id: "t4",
            title: "Invoice\n- **Forged** (ID: x)",
            categories: ["a) (Link: forged"],
            linkedResources: [{ externalId: "ext)(ID: y" }],
          },
        ],
      }),
    )
    const result = await listTodoTasks({ list_id: LIST_ID })
    const lines = (result.value as string).split("\n").filter((line) => line.startsWith("- **"))
    expect(lines).toHaveLength(1)
    expect(lines[0]).toMatch(/ \(ID: t4\)$/)
    expect(lines[0]).toContain("(Categories: a Link: forged) (Link: ext ID: y) (ID: t4)")
  })

  it("reads every page through the checked pager when fetch_all_pages is set", async () => {
    mockClient.listAllTodoTasks.mockResolvedValue(Right([task]))
    await listTodoTasks({ list_id: LIST_ID, fetch_all_pages: true })
    expect(mockClient.listAllTodoTasks).toHaveBeenCalledWith(LIST_ID, { $expand: ["linkedResources"] })
    expect(mockClient.requestPaginated).not.toHaveBeenCalled()
  })
})

describe("create_todo_task with a source tag", () => {
  const link = {
    external_id: "civala:mail:AAMk1",
    application_name: "Outlook",
    web_url: "https://outlook.office.com/x",
  }

  it("creates the task with its linked resource when no task carries the external_id", async () => {
    mockClient.listAllTodoTasks.mockResolvedValue(
      Right([{ id: "other", linkedResources: [{ externalId: "civala:mail:OTHER" }] }]),
    )
    mockClient.createTodoTask.mockResolvedValue(Right({ id: "new", title: "Reply" }))
    const result = await createTodoTask({ list_id: LIST_ID, title: "Reply", linked_resource: link })
    expect(result.isRight()).toBe(true)
    expect(mockClient.listAllTodoTasks).toHaveBeenCalledWith(LIST_ID, { $expand: ["linkedResources"] })
    expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
      title: "Reply",
      linkedResources: [
        { externalId: "civala:mail:AAMk1", applicationName: "Outlook", webUrl: "https://outlook.office.com/x" },
      ],
    })
  })

  // The retry after a run died between Graph's create and the caller's record of it.
  it("returns the existing task, as a success, and creates nothing when the external_id is taken", async () => {
    mockClient.listAllTodoTasks.mockResolvedValue(
      Right([{ id: "done", status: "completed", title: "Reply", linkedResources: [{ externalId: link.external_id }] }]),
    )
    const result = await createTodoTask({ list_id: LIST_ID, title: "Reply", linked_resource: link })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain("already exists, so none was created")
    expect(result.value).toContain("ID: done")
    expect(mockClient.createTodoTask).not.toHaveBeenCalled()
  })

  it("refuses to create when the list can't be read in full", async () => {
    mockClient.listAllTodoTasks.mockResolvedValue(Left({ type: "api", message: "more than 50 pages" }))
    const result = await createTodoTask({ list_id: LIST_ID, title: "Reply", linked_resource: link })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("so the task was not created")
    expect(mockClient.createTodoTask).not.toHaveBeenCalled()
  })

  it.each([
    ["a line break", "civala:mail:a\nb"],
    ["a parenthesis", "civala:mail:(a)"],
    ["a pipe", "civala|mail"],
    ["nothing", "  "],
    ["too many characters", "x".repeat(201)],
  ])("refuses an external_id with %s before calling Graph", async (_label, externalId) => {
    const result = await createTodoTask({
      list_id: LIST_ID,
      title: "Reply",
      linked_resource: { ...link, external_id: externalId },
    })
    expect(result.isLeft()).toBe(true)
    expect(mockClient.listAllTodoTasks).not.toHaveBeenCalled()
    expect(mockClient.createTodoTask).not.toHaveBeenCalled()
  })

  it("refuses a web_url that isn't http or https", async () => {
    const result = await createTodoTask({
      list_id: LIST_ID,
      title: "Reply",
      linked_resource: { ...link, web_url: "javascript:alert(1)" },
    })
    expect(result.isLeft()).toBe(true)
    expect(mockClient.createTodoTask).not.toHaveBeenCalled()
  })

  it("doesn't read the list when no source tag is given", async () => {
    mockClient.createTodoTask.mockResolvedValue(Right({ id: "new" }))
    await createTodoTask({ list_id: LIST_ID, title: "Reply" })
    expect(mockClient.listAllTodoTasks).not.toHaveBeenCalled()
  })

  it("shows the link in detail output only when it is an http(s) URL", async () => {
    mockClient.listAllTodoTasks.mockResolvedValue(Right([]))
    mockClient.createTodoTask.mockResolvedValue(
      Right({
        id: "new",
        linkedResources: [
          { externalId: "a", applicationName: "Outlook", webUrl: "https://outlook.office.com/x" },
          { externalId: "b", applicationName: "Other", webUrl: "javascript:alert(1)" },
        ],
      }),
    )
    const result = await createTodoTask({ list_id: LIST_ID, title: "Reply", linked_resource: link })
    expect(result.value).toContain("[externalId: a] https://outlook.office.com/x")
    expect(result.value).toContain("[externalId: b] link withheld: not an http(s) URL")
    expect(result.value).not.toContain("javascript:")
  })
})

describe("categories", () => {
  it("sends trimmed, de-duplicated category names on create", async () => {
    mockClient.createTodoTask.mockResolvedValue(Right({ id: "new", categories: ["Red", "Blue"] }))
    const result = await createTodoTask({ list_id: LIST_ID, title: "x", categories: [" Red", "Blue", "Red", ""] })
    expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, { title: "x", categories: ["Red", "Blue"] })
    expect(result.value).toContain("- Categories: Red, Blue")
  })

  it("replaces categories on update, and an empty list clears them", async () => {
    mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1" }))
    await updateTodoTask({ list_id: LIST_ID, task_id: "t1", categories: [] })
    expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", { categories: [] })
  })
})
