import { randomUUID } from "node:crypto"

import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphGroup, GraphPlan, GraphPlannerTask, ODataResponse } from "../types"
import { formatBucketList, formatPlanList, formatPlannerTaskDetail, formatPlannerTaskList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

// /me/planner/plans only returns plans the user was explicitly added to; group-owned plans reachable
// via membership are omitted — a skill calling this would silently see an empty board. So fan out over
// the user's groups and merge with /me. Per-source failures (a group without Planner, a transient 4xx)
// are skipped rather than fatal; only an all-sources failure surfaces an error, so one bad group can't
// blank the whole board.
export const listPlans = async (): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const groupsResult = await client.listMyGroups()
  const groups = groupsResult.fold<ReadonlyArray<GraphGroup>>(
    () => [],
    (r) => r.value,
  )

  const sources = await Promise.all([client.listPlans(), ...groups.map((g) => client.listGroupPlans(g.id))])

  if (!sources.some((r) => r.isRight())) {
    const failed = [groupsResult, ...sources].find((r) => r.isLeft())
    const message = failed ? (failed.value as { message: string }).message : "unknown error"
    return Left(new UserError(`Failed to list plans: ${message}`))
  }

  const plans = sources.flatMap((r) =>
    r.fold<ReadonlyArray<GraphPlan>>(
      () => [],
      (resp) => resp.value,
    ),
  )
  const deduped = [...new Map(plans.map((p) => [p.id, p])).values()]
  return Right(formatPlanList(deduped))
}

// Buckets are the columns a plan's tasks sort into. create_planner_task takes a bucket_id but had no
// source for one; these expose the list/create so a valid id can be supplied.
export const listPlannerBuckets = async (params: { plan_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listPlannerBuckets(params.plan_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list buckets: ${error.message}`))
    .map((response) => formatBucketList(response.value))
}

export const createPlannerBucket = async (params: {
  plan_id: string
  name: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.createPlannerBucket({ name: params.name, planId: params.plan_id, orderHint: " !" })
  return result
    .mapLeft((error) => new UserError(`Failed to create bucket: ${error.message}`))
    .map((b) => `Bucket created.\n\n- **${b.name ?? params.name}** (ID: ${b.id})`)
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

// The task-level ETag is separate from the details ETag. `etag` is optional: pass one to enforce
// optimistic concurrency (a 412 then means someone else edited first — surfaced, not retried); omit it
// and the current ETag is fetched for you, retrying once on a 412 — consistent with
// update_planner_task_details, which always auto-fetches.
export const updatePlannerTask = async (params: {
  task_id: string
  etag?: string
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

  const resolveEtag = async (): Promise<Either<UserError, string>> => {
    if (params.etag) return Right(params.etag)
    const task = await client.getPlannerTask(params.task_id)
    return task.fold<Either<UserError, string>>(
      (error) => Left(new UserError(`Failed to read task: ${error.message}`)),
      (value) => Right((value as unknown as { "@odata.etag": string })["@odata.etag"]),
    )
  }
  const attempt = (etag: string) => client.updatePlannerTask(params.task_id, updates, etag)

  // Planner's task PATCH returns 204 with no body, so there's nothing to format — summarize the
  // requested changes instead of rendering an empty task (which read as "Untitled / ID: undefined").
  const summary = [
    params.title !== undefined ? `title="${params.title}"` : undefined,
    params.percent_complete !== undefined ? `percentComplete=${params.percent_complete}` : undefined,
    params.due_date !== undefined ? `due=${params.due_date}` : undefined,
    params.priority !== undefined ? `priority=${params.priority}` : undefined,
  ]
    .filter((s): s is string => s !== undefined)
    .join(", ")
  const successMessage = summary ? `Task updated (${summary}).` : "Task updated."
  // Surface the HTTP status so callers can distinguish a 412 conflict from a 400/permission error
  // without string-matching the message.
  const failure = (e: { status?: number; message: string }) =>
    new UserError(`Failed to update task${e.status ? ` (HTTP ${e.status})` : ""}: ${e.message}`)

  const firstEtag = await resolveEtag()
  if (firstEtag.isLeft()) return firstEtag
  const firstResult = await attempt(firstEtag.value as string)
  if (firstResult.isRight()) return Right(successMessage)

  // Retry a 412 only when we fetched the ETag; if the caller pinned one, respect their concurrency guard.
  const firstError = firstResult.value as { status?: number; message: string }
  if (firstError.status !== 412 || params.etag) return Left(failure(firstError))

  const retryEtag = await resolveEtag()
  if (retryEtag.isLeft()) return retryEtag
  return (await attempt(retryEtag.value as string)).fold<Either<UserError, string>>(
    (error) => Left(failure(error as { status?: number; message: string })),
    () => Right(successMessage),
  )
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

// The current details object as far as we address it: checklist keyed by GUID, references keyed by the
// encoded external URL.
type DetailsSnapshot = {
  readonly "@odata.etag": string
  readonly checklist?: Record<string, { title?: string; isChecked?: boolean }>
  readonly references?: Record<string, { alias?: string; type?: string }>
}

const checklistItem = (title: string, isChecked: boolean) => ({
  "@odata.type": "#microsoft.graph.plannerChecklistItem",
  title,
  isChecked,
})
const referenceItem = (alias: string, type: string) => ({
  "@odata.type": "#microsoft.graph.plannerExternalReference",
  alias,
  type,
})

// Task description / checklist / references live on a separate task-details object with its own If-Match
// ETag. checklist and references are open-typed maps: a key present in the PATCH adds/replaces it, a key
// set to null deletes it, an absent key is untouched. This reads the current details (for the ETag and
// to merge/validate edits), builds ONE PATCH covering adds + updates + removes, and retries once on 412.
// Updates are read-modify-write (omitted fields keep their current value) so a partial edit can't blank
// the rest of an item. Removing/updating a key Graph doesn't have is a silent no-op there, so those are
// reported as "skipped" rather than falsely succeeding.
export const updatePlannerTaskDetails = async (params: {
  task_id: string
  description?: string
  preview_type?: "automatic" | "noPreview" | "checklist" | "description" | "reference"
  add_checklist?: ReadonlyArray<{ title: string; isChecked?: boolean }>
  update_checklist?: ReadonlyArray<{ id: string; title?: string; isChecked?: boolean }>
  remove_checklist?: ReadonlyArray<string>
  add_references?: ReadonlyArray<{ url: string; alias?: string }>
  update_references?: ReadonlyArray<{ url: string; alias?: string }>
  remove_references?: ReadonlyArray<string>
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const buildPatch = (
    current: DetailsSnapshot,
  ): { details: Record<string, unknown>; skipped: ReadonlyArray<string> } => {
    const skipped: string[] = []
    const details: Record<string, unknown> = {}
    if (params.description !== undefined) details.description = params.description
    if (params.preview_type) details.previewType = params.preview_type

    const existingChecklist = current.checklist ?? {}
    const checklist: Record<string, unknown> = {}
    const clAdds = params.add_checklist ?? []
    clAdds.forEach((c) => {
      checklist[randomUUID()] = checklistItem(c.title, c.isChecked ?? false)
    })
    const clUpdates = params.update_checklist ?? []
    clUpdates.forEach((u) => {
      const cur = existingChecklist[u.id]
      if (!cur) {
        skipped.push(`checklist item ${u.id}`)
        return
      }
      checklist[u.id] = checklistItem(u.title ?? cur.title ?? "", u.isChecked ?? cur.isChecked ?? false)
    })
    const clRemoves = params.remove_checklist ?? []
    clRemoves.forEach((id) => {
      if (!existingChecklist[id]) {
        skipped.push(`checklist item ${id}`)
        return
      }
      checklist[id] = null
    })
    if (Object.keys(checklist).length > 0) details.checklist = checklist

    const existingRefs = current.references ?? {}
    const references: Record<string, unknown> = {}
    const refAdds = params.add_references ?? []
    refAdds.forEach((r) => {
      references[encodeRefKey(r.url)] = referenceItem(r.alias ?? r.url, "Other")
    })
    const refUpdates = params.update_references ?? []
    refUpdates.forEach((r) => {
      const key = encodeRefKey(r.url)
      const cur = existingRefs[key]
      if (!cur) {
        skipped.push(`reference ${r.url}`)
        return
      }
      references[key] = referenceItem(r.alias ?? cur.alias ?? r.url, cur.type ?? "Other")
    })
    const refRemoves = params.remove_references ?? []
    refRemoves.forEach((url) => {
      const key = encodeRefKey(url)
      if (!existingRefs[key]) {
        skipped.push(`reference ${url}`)
        return
      }
      references[key] = null
    })
    if (Object.keys(references).length > 0) details.references = references

    return { details, skipped }
  }

  // One attempt: read current details (ETag + existing keys), build the merged PATCH, write it. Planner
  // requires a concrete If-Match (wildcard "*" isn't honored). retriable is true only for a 412.
  const runOnce = async (): Promise<Either<{ retriable: boolean; error: UserError }, string>> => {
    const currentResult = await client.getPlannerTaskDetails(params.task_id)
    if (currentResult.isLeft()) {
      const { message } = currentResult.value as { message: string }
      return Left({ retriable: false, error: new UserError(`Failed to read task details: ${message}`) })
    }
    const current = currentResult.value as unknown as DetailsSnapshot
    const { details, skipped } = buildPatch(current)

    const patchResult = await client.updatePlannerTaskDetails(params.task_id, details, current["@odata.etag"])
    return patchResult.fold<Either<{ retriable: boolean; error: UserError }, string>>(
      (error) => {
        const e = error as { status?: number; message: string }
        const wrapped = new UserError(
          `Failed to update task details${e.status ? ` (HTTP ${e.status})` : ""}: ${e.message}`,
        )
        return Left({ retriable: e.status === 412, error: wrapped })
      },
      () =>
        Right(
          skipped.length > 0
            ? `Task details updated. Skipped (not found): ${skipped.join(", ")}.`
            : "Task details updated.",
        ),
    )
  }

  const first = await runOnce()
  if (first.isRight()) return Right(first.value as string)
  const { retriable, error } = first.value as { retriable: boolean; error: UserError }
  if (!retriable) return Left(error)

  // Concurrent edit invalidated the ETag (412): re-read and retry exactly once.
  return (await runOnce()).fold<Either<UserError, string>>(
    (l) => Left((l as { error: UserError }).error),
    (message) => Right(message as string),
  )
}
