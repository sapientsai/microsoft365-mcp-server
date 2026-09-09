import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphTodoList, GraphTodoTask, ODataResponse } from "../types"
import { formatTodoListList, formatTodoTaskDetail, formatTodoTaskList } from "../utils/formatters"
import type { RecurrenceInput } from "../utils/recurrence"
import { buildRecurrence } from "../utils/recurrence"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listTodoLists = async (params?: { fetch_all_pages?: boolean }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphTodoList>("/me/todo/lists")
    return result
      .mapLeft((error) => new UserError(`Failed to list To Do lists: ${error.message}`))
      .map((items) => formatTodoListList(items))
  }

  const result = await client.listTodoLists()
  return result
    .mapLeft((error) => new UserError(`Failed to list To Do lists: ${error.message}`))
    .map((response) => formatTodoListList((response as ODataResponse<never>).value))
}

export const listTodoTasks = async (params: {
  list_id: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphTodoTask>(`/me/todo/lists/${params.list_id}/tasks`)
    return result
      .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
      .map((items) => formatTodoTaskList(items))
  }

  const result = await client.listTodoTasks(params.list_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
    .map((response) => formatTodoTaskList((response as ODataResponse<never>).value))
}

export const createTodoTask = async (params: {
  list_id: string
  title: string
  body?: string
  due_date?: string
  importance?: string
  recurrence?: RecurrenceInput
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const task: Record<string, unknown> = { title: params.title }
  if (params.body) task.body = { contentType: "text", content: params.body }
  if (params.due_date) task.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
  if (params.importance) task.importance = params.importance

  if (params.recurrence) {
    // To Do rolls a recurring task forward from its due date, so a recurrence with
    // no due date creates a task that repeats but never appears in Today. Graph
    // accepts it silently, which makes this worth refusing here rather than
    // shipping a task the user cannot see.
    if (!params.due_date) {
      return Left(new UserError("A recurring task needs a due_date — To Do repeats a task from its due date."))
    }

    const recurrence = buildRecurrence(params.recurrence, params.due_date)
    if (typeof recurrence === "string") return Left(new UserError(recurrence))
    task.recurrence = recurrence
  }

  const result = await client.createTodoTask(params.list_id, task)
  return result
    .mapLeft((error) => new UserError(`Failed to create task: ${error.message}`))
    .map((t) => `Task created.\n\n${formatTodoTaskDetail(t)}`)
}

export const updateTodoTask = async (params: {
  list_id: string
  task_id: string
  title?: string
  status?: string
  due_date?: string
  importance?: string
  body?: string
  recurrence?: RecurrenceInput
  clear_recurrence?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.recurrence && params.clear_recurrence) {
    return Left(new UserError("Pass either recurrence or clear_recurrence, not both."))
  }

  const updates: Record<string, unknown> = {}
  if (params.title) updates.title = params.title
  if (params.status) updates.status = params.status
  if (params.due_date) updates.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
  if (params.importance) updates.importance = params.importance
  if (params.body) updates.body = { contentType: "text", content: params.body }

  // Graph clears a recurrence by an explicit null; omitting the property leaves the
  // existing pattern in place, so the two cases cannot share a code path.
  if (params.clear_recurrence) updates.recurrence = null

  if (params.recurrence) {
    const recurrence = buildRecurrence(params.recurrence, params.due_date)
    if (typeof recurrence === "string") return Left(new UserError(recurrence))
    updates.recurrence = recurrence
  }

  const result = await client.updateTodoTask(params.list_id, params.task_id, updates)
  return result
    .mapLeft((error) => new UserError(`Failed to update task: ${error.message}`))
    .map((t) => `Task updated.\n\n${formatTodoTaskDetail(t)}`)
}
