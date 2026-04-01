import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphTodoList, GraphTodoTask, ODataResponse } from "../types"
import { formatTodoListList, formatTodoTaskDetail, formatTodoTaskList } from "../utils/formatters"

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
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const task: Record<string, unknown> = { title: params.title }
  if (params.body) task.body = { contentType: "Text", content: params.body }
  if (params.due_date) task.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
  if (params.importance) task.importance = params.importance

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
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const updates: Record<string, unknown> = {}
  if (params.title) updates.title = params.title
  if (params.status) updates.status = params.status
  if (params.due_date) updates.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
  if (params.importance) updates.importance = params.importance
  if (params.body) updates.body = { contentType: "Text", content: params.body }

  const result = await client.updateTodoTask(params.list_id, params.task_id, updates)
  return result
    .mapLeft((error) => new UserError(`Failed to update task: ${error.message}`))
    .map((t) => `Task updated.\n\n${formatTodoTaskDetail(t)}`)
}
