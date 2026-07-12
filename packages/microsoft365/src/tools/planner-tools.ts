import { randomUUID } from "node:crypto"

import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphPlan, GraphPlannerTask, ODataResponse } from "../types"
import { formatPlanList, formatPlannerTaskDetail, formatPlannerTaskList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listPlans = async (params?: { fetch_all_pages?: boolean }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params?.fetch_all_pages) {
    const result = await client.requestPaginated<GraphPlan>("/me/planner/plans")
    return result
      .mapLeft((error) => new UserError(`Failed to list plans: ${error.message}`))
      .map((items) => formatPlanList(items))
  }

  const result = await client.listPlans()
  return result
    .mapLeft((error) => new UserError(`Failed to list plans: ${error.message}`))
    .map((response) => formatPlanList((response as ODataResponse<never>).value))
}

export const listPlannerTasks = async (params: {
  plan_id: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphPlannerTask>(`/planner/plans/${params.plan_id}/tasks`)
    return result
      .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
      .map((items) => formatPlannerTaskList(items))
  }

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

// Planner reference keys are the external URL used as an OData open-type property name. Only the
// characters illegal in a property name get escaped; "/" and the rest of the URL stay intact so
// Planner still parses it as a URI (matches Microsoft's documented key shape, e.g.
// http%3A//developer%2Emicrosoft%2Ecom). encodeURIComponent is WRONG here — it escapes the slashes
// too, and Planner then rejects the key ("The Authority/Host could not be parsed"). Escape "%" first
// so we don't double-encode. Query-string chars (?, &, =) aren't in Planner's forbidden set — left
// as-is. Verified live against Graph 2026-07-12.
const encodeRefKey = (url: string): string =>
  url.replace(/%/g, "%25").replace(/\./g, "%2E").replace(/:/g, "%3A").replace(/@/g, "%40").replace(/#/g, "%23")

// Task description / checklist / references live on a separate task-details object that requires its
// own If-Match ETag. This reads that ETag, PATCHes the details, and retries once on a 412 (a
// concurrent edit invalidating the ETag between the read and the write).
export const updatePlannerTaskDetails = async (params: {
  task_id: string
  description?: string
  preview_type?: "automatic" | "noPreview" | "checklist" | "description" | "reference"
  add_checklist?: ReadonlyArray<{ title: string; isChecked?: boolean }>
  add_references?: ReadonlyArray<{ url: string; alias?: string }>
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  // Build the etag-independent payload once. checklist / references are open-typed maps: including a
  // key ADDS/updates it (new items are added, not replaced). Removing an item is a follow-up PATCH of
  // its key to null — out of scope here.
  const details: Record<string, unknown> = {}
  if (params.description !== undefined) details.description = params.description
  if (params.preview_type) details.previewType = params.preview_type
  if (params.add_checklist) {
    details.checklist = Object.fromEntries(
      params.add_checklist.map((c) => [
        randomUUID(),
        { "@odata.type": "#microsoft.graph.plannerChecklistItem", title: c.title, isChecked: c.isChecked ?? false },
      ]),
    )
  }
  if (params.add_references) {
    details.references = Object.fromEntries(
      params.add_references.map((r) => [
        encodeRefKey(r.url),
        { "@odata.type": "#microsoft.graph.plannerExternalReference", alias: r.alias ?? r.url, type: "Other" },
      ]),
    )
  }

  // Planner requires a concrete If-Match on the details object (wildcard "*" is not honored), so each
  // attempt reads the current ETag immediately before PATCHing.
  const readEtag = async (): Promise<Either<UserError, string>> => {
    const current = await client.getPlannerTaskDetails(params.task_id)
    return current.fold<Either<UserError, string>>(
      (error) => Left(new UserError(`Failed to read task details: ${error.message}`)),
      (value) => Right(value["@odata.etag"]),
    )
  }
  const patch = (etag: string) => client.updatePlannerTaskDetails(params.task_id, details, etag)

  const firstEtag = await readEtag()
  if (firstEtag.isLeft()) return firstEtag
  const firstResult = await patch(firstEtag.value as string)
  if (firstResult.isRight()) return Right("Task details updated.")

  // A concurrent edit invalidated the ETag between our read and PATCH (HTTP 412). Re-read once and
  // retry so an unattended loop doesn't lose a write to a benign race; surface any other error as-is.
  const firstError = firstResult.value as { status?: number; message: string }
  if (firstError.status !== 412) {
    return Left(new UserError(`Failed to update task details: ${firstError.message}`))
  }

  const retryEtag = await readEtag()
  if (retryEtag.isLeft()) return retryEtag
  return (await patch(retryEtag.value as string))
    .mapLeft((error) => new UserError(`Failed to update task details: ${error.message}`))
    .map(() => "Task details updated.")
}
