import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { ODataResponse } from "../types"
import { formatPlanList, formatPlannerTaskDetail, formatPlannerTaskList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listPlans = async (): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listPlans()
  return result
    .mapLeft((error) => new UserError(`Failed to list plans: ${error.message}`))
    .map((response) => formatPlanList((response as ODataResponse<never>).value))
}

export const listPlannerTasks = async (params: { plan_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listPlannerTasks(params.plan_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
    .map((response) => formatPlannerTaskList((response as ODataResponse<never>).value))
}

export const getPlannerTask = async (params: { task_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getPlannerTask(params.task_id)
  return result.mapLeft((error) => new UserError(`Failed to get task: ${error.message}`)).map(formatPlannerTaskDetail)
}

export const createPlannerTask = async (params: {
  plan_id: string
  title: string
  bucket_id?: string
  due_date?: string
  assignments?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const task: Record<string, unknown> = {
    planId: params.plan_id,
    title: params.title,
  }
  if (params.bucket_id) task.bucketId = params.bucket_id
  if (params.due_date) task.dueDateTime = params.due_date
  if (params.assignments) {
    const assignees: Record<string, { "@odata.type": string; orderHint: string }> = {}
    params.assignments.split(",").forEach((userId) => {
      assignees[userId.trim()] = { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" }
    })
    task.assignments = assignees
  }

  const result = await client.createPlannerTask(task)
  return result
    .mapLeft((error) => new UserError(`Failed to create task: ${error.message}`))
    .map((t) => `Task created.\n\n${formatPlannerTaskDetail(t)}`)
}

export const updatePlannerTask = async (params: {
  task_id: string
  etag: string
  title?: string
  percent_complete?: number
  due_date?: string
  priority?: number
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const updates: Record<string, unknown> = {}
  if (params.title) updates.title = params.title
  if (params.percent_complete !== undefined) updates.percentComplete = params.percent_complete
  if (params.due_date) updates.dueDateTime = params.due_date
  if (params.priority !== undefined) updates.priority = params.priority

  const result = await client.updatePlannerTask(params.task_id, updates, params.etag)
  return result
    .mapLeft((error) => new UserError(`Failed to update task: ${error.message}`))
    .map((t) => `Task updated.\n\n${formatPlannerTaskDetail(t)}`)
}
