import { UserError } from "fastmcp"
import { Option } from "functype"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient, type TodoDeltaRecord, type TodoDeltaResult } from "../client/graph-client"
import type { GraphApiError, GraphTodoList, GraphTodoTask, ODataResponse } from "../types"
import { isDeltaToken } from "../utils/delta-token"
import {
  extensionJson,
  formatTodoListList,
  formatTodoTaskChanges,
  formatTodoTaskDetail,
  formatTodoTaskList,
} from "../utils/formatters"
import { displayableLinkUrl } from "../utils/link-url"
import type { RecurrenceInput } from "../utils/recurrence"
import { buildRecurrence, describeRecurrence, toDateOnly } from "../utils/recurrence"
import {
  extensionExpand,
  extensionFields,
  extensionFieldsProblem,
  type ExtensionInput,
  extensionNameProblem,
  extensionProblem,
  type ExtensionValue,
  toExtensionBody,
  toExtensionPatch,
} from "../utils/todo-extension"

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

// Graph accepts a duplicate displayName and returns a second list with the same name,
// which is indistinguishable from the first in every client. Existing lists are read
// first and an exact match refused, because the usual cause is a caller that did not
// know the list was already there — and a duplicate is worse than the error, since
// tasks then scatter across two lists nobody can tell apart. Every page is read, and a
// failed read refuses rather than creating blind, since the guard is only as good as
// the list it checks against.
export const createTodoList = async (params: { display_name: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const name = params.display_name.trim()
  if (name.length === 0) return Left(new UserError("A To Do list needs a display_name."))

  const existing = await client.requestPaginated<GraphTodoList>("/me/todo/lists")
  if (existing.isLeft()) {
    return Left(
      new UserError(
        `Failed to read existing To Do lists, so the list was not created: ${(existing.value as GraphApiError).message}`,
      ),
    )
  }
  const lists = existing.value as GraphTodoList[]
  const clash = lists.find((list) => list.displayName?.trim().toLowerCase() === name.toLowerCase())
  if (clash) {
    return Left(
      new UserError(
        `A To Do list named "${clash.displayName}" already exists (ID: ${clash.id}). ` +
          "Use it, or choose a different name — two lists with one name cannot be told apart.",
      ),
    )
  }

  const result = await client.createTodoList(name)
  return result
    .mapLeft((error) => new UserError(`Failed to create To Do list: ${error.message}`))
    .map(
      (list) =>
        `Created To Do list **${list.displayName ?? name}** (ID: ${list.id}).\n\n` +
        "Graph cannot share a list — sharing is only available in the To Do app, and a new list " +
        "is visible to nobody else until someone shares it there.",
    )
}

export type LinkedResourceInput = {
  readonly external_id: string
  readonly application_name: string
  readonly web_url?: string
  readonly display_name?: string
}

// externalId is the key a caller looks a task up by, so it has to come back exactly as sent. Output
// can't clean it without breaking that match, so an unsafe one is refused here instead.
const EXTERNAL_ID_MAX = 200

const linkedResourceProblem = (link: LinkedResourceInput): string | undefined => {
  if (link.external_id.trim() === "") return "linked_resource.external_id must not be empty."
  if (link.external_id.length > EXTERNAL_ID_MAX) {
    return `linked_resource.external_id is longer than ${EXTERNAL_ID_MAX} characters.`
  }
  // A comma would make two linked ids on one line ambiguous, since the list line joins them with ", ".
  if (/[\r\n\v\f\u0085\u2028\u2029|(),]/.test(link.external_id)) {
    return "linked_resource.external_id must not contain a line break, |, (, ) or a comma."
  }
  if (link.application_name.trim() === "") return "linked_resource.application_name must not be empty."
  if (link.web_url !== undefined && displayableLinkUrl(link.web_url) === undefined) {
    return "linked_resource.web_url must be an http: or https: URL."
  }
  return undefined
}

// Graph's linkedResource docs list displayName as required, so it falls back to the app's name.
const toLinkedResource = (link: LinkedResourceInput): Record<string, string> => ({
  externalId: link.external_id,
  applicationName: link.application_name,
  displayName: link.display_name ?? link.application_name,
  ...(link.web_url === undefined ? {} : { webUrl: link.web_url }),
})

const cleanCategories = (categories: ReadonlyArray<string>): ReadonlyArray<string> => [
  ...new Set(categories.map((c) => c.trim()).filter((c) => c !== "")),
]

// A task's source lives in linkedResources, a relationship Graph returns only when expanded. With an
// extension name, that one extension is expanded too, so its fields can print on each line.
const taskExpand = (extensionName?: string) => ({
  $expand: extensionName === undefined ? ["linkedResources"] : ["linkedResources", extensionExpand(extensionName)],
})

export const listTodoTasks = async (params: {
  list_id: string
  fetch_all_pages?: boolean
  extension_name?: string
  delta?: boolean
  delta_token?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const nameProblem = params.extension_name === undefined ? undefined : extensionNameProblem(params.extension_name)
  if (nameProblem) return Left(new UserError(nameProblem))
  if (params.delta === true || params.delta_token !== undefined) return listTodoTaskChanges(client, params)
  const expand = taskExpand(params.extension_name)

  if (params.fetch_all_pages) {
    const result = await client.listAllTodoTasks(params.list_id, expand)
    return result
      .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
      .map((items) => formatTodoTaskList(items, params.extension_name))
  }

  const result = await client.listTodoTasks(params.list_id, expand)
  return result
    .mapLeft((error) => new UserError(`Failed to list tasks: ${error.message}`))
    .map((response) => formatTodoTaskList((response as ODataResponse<GraphTodoTask>).value, params.extension_name))
}

export const createTodoTask = async (params: {
  list_id: string
  title: string
  body?: string
  due_date?: string
  importance?: string
  recurrence?: RecurrenceInput
  categories?: ReadonlyArray<string>
  linked_resource?: LinkedResourceInput
  extension?: ExtensionInput
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const linkProblem = params.linked_resource ? linkedResourceProblem(params.linked_resource) : undefined
  if (linkProblem) return Left(new UserError(linkProblem))
  const extProblem = params.extension ? extensionProblem(params.extension) : undefined
  if (extProblem) return Left(new UserError(extProblem))

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

  if (params.categories) task.categories = cleanCategories(params.categories)
  // Created inside the task's own POST (checked live 2026-10-30), so there's no second call to fail.
  if (params.extension) task.extensions = [toExtensionBody(params.extension.name, params.extension.fields)]

  // A caller that died after Graph created the task, but before it recorded that, retries the
  // create. With an external_id, the retry finds the task it already made instead of making a
  // second one. Graph can't filter on linkedResources ("CollectionNavigationNode operation is not
  // supported", checked live 2026-10-30), so every page is read and matched here, completed tasks
  // included. A list that can't be read in full refuses the create rather than risk the duplicate.
  if (params.linked_resource) {
    const externalId = params.linked_resource.external_id
    const existing = await client.listAllTodoTasks(params.list_id, taskExpand(params.extension?.name))
    if (existing.isLeft()) {
      return Left(
        new UserError(
          `Failed to read the list's tasks, so the task was not created: ${(existing.value as GraphApiError).message}`,
        ),
      )
    }
    const match = (existing.value as ReadonlyArray<GraphTodoTask>).find((t) =>
      (t.linkedResources ?? []).some((r) => r.externalId === externalId),
    )
    // Already there is the outcome the caller wanted, so it is a success, not an error: see
    // AGENTS.md, Results, on an idempotent repeat.
    if (match) {
      return Right(
        `A task with this external_id already exists, so none was created.\n\n${formatTodoTaskDetail(match, params.extension?.name)}`,
      )
    }
    task.linkedResources = [toLinkedResource(params.linked_resource)]
  }

  const result = await client.createTodoTask(params.list_id, task)
  return result
    .mapLeft((error) => new UserError(`Failed to create task: ${error.message}`))
    .map((t) => `Task created.\n\n${formatTodoTaskDetail(t, params.extension?.name)}`)
}

export const deleteTodoTask = async (params: {
  list_id: string
  task_id: string
  force?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  // Read before deleting. A task id addresses the whole recurring series, not one
  // occurrence, so deleting one ends the series — and Graph has no restore. Microsoft
  // documents a manual recovery from Outlook's Deleted Items, unverified for API
  // deletes, so this is the last reliable chance to say so.
  const existing = await client.getTodoTask(params.list_id, params.task_id)
  if (existing.isLeft()) {
    return Left(new UserError(`Failed to read the task before deleting: ${(existing.value as GraphApiError).message}`))
  }

  const task = existing.value as GraphTodoTask
  const title = task.title ?? "Untitled"

  if (task.recurrence && !params.force) {
    return Left(
      new UserError(
        `"${title}" repeats ${describeRecurrence(task.recurrence)}. Deleting it ends the whole series, ` +
          `not just this occurrence, and this tool can't undo it. Pass force: true to go ahead, or use ` +
          `clear_recurrence on update_todo_task to keep the task but stop it repeating.`,
      ),
    )
  }

  const result = await client.deleteTodoTask(params.list_id, params.task_id)
  return result
    .mapLeft((error) => new UserError(`Failed to delete task: ${error.message}`))
    .map(() =>
      task.recurrence
        ? `Deleted "${title}" and its whole recurring series (was ${describeRecurrence(task.recurrence)}).`
        : `Deleted "${title}".`,
    )
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
  categories?: ReadonlyArray<string>
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.recurrence && params.clear_recurrence) {
    return Left(new UserError("Pass either recurrence or clear_recurrence, not both."))
  }

  // On a repeating task both dates set where the series starts, so two different
  // ones leave no right answer. Refuse rather than let one silently win.
  const startDate = params.recurrence?.start_date
  if (params.due_date && startDate && toDateOnly(params.due_date) !== toDateOnly(startDate)) {
    return Left(
      new UserError(
        `due_date (${toDateOnly(params.due_date)}) and recurrence.start_date (${toDateOnly(startDate)}) differ. ` +
          "On a repeating task both set when it starts, so pass one, or the same date in both.",
      ),
    )
  }

  const updates: Record<string, unknown> = {}
  if (params.title) updates.title = params.title
  if (params.status) updates.status = params.status
  if (params.importance) updates.importance = params.importance
  if (params.body) updates.body = { contentType: "text", content: params.body }
  // Replaces the task's categories; an empty list clears them.
  if (params.categories) updates.categories = cleanCategories(params.categories)

  // Graph clears a recurrence by an explicit null; omitting the property leaves the
  // existing pattern in place, so the two cases cannot share a code path.
  if (params.clear_recurrence) updates.recurrence = null

  if (params.recurrence) {
    const recurrence = buildRecurrence(params.recurrence, params.due_date)
    if (typeof recurrence === "string") return Left(new UserError(recurrence))
    updates.recurrence = recurrence
  }

  // A dueDateTime PATCH on a repeating task does not move it. Graph treats it as
  // completing the current occurrence: the series rolls forward one period and a
  // second task appears on the requested date — with or without a recurrence in the
  // same PATCH (checked live 2026-10-28). Moving the recurrence's startDate moves the
  // due date instead, on the same task. So on a repeating task the new date travels
  // as the startDate and dueDateTime is never sent.
  if (params.due_date && !params.clear_recurrence) {
    if (!params.recurrence) {
      const existing = await client.getTodoTask(params.list_id, params.task_id)
      if (existing.isLeft()) {
        return Left(
          new UserError(`Failed to read the task before updating: ${(existing.value as GraphApiError).message}`),
        )
      }
      const { recurrence } = existing.value as GraphTodoTask
      if (recurrence) {
        const newStart = toDateOnly(params.due_date)
        // Graph echoes an endDate of 0001-01-01 on noEnd ranges, so only an endDate
        // range has an end worth checking.
        const { range } = recurrence
        if (range.type === "endDate" && range.endDate !== undefined && range.endDate < newStart) {
          return Left(
            new UserError(
              `This task stops repeating on ${range.endDate}, before ${newStart}. Pass recurrence with a later ` +
                "end_date to move it, or clear_recurrence to make it one-off.",
            ),
          )
        }
        updates.recurrence = { ...recurrence, range: { ...range, startDate: newStart } }
      } else {
        updates.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
      }
    }
  } else if (params.due_date) {
    updates.dueDateTime = { dateTime: params.due_date, timeZone: "UTC" }
  }

  const result = await client.updateTodoTask(params.list_id, params.task_id, updates)
  return result
    .mapLeft((error) => new UserError(`Failed to update task: ${error.message}`))
    .map((t) => {
      // Completing a repeating task already reports the next occurrence, so a
      // second sentence about where the move landed could only contradict it.
      const reschedule = params.status === "completed" ? "" : describeReschedule(params.due_date, updates, t)
      const summary = [describeUpdate(params.status, t), reschedule].filter((line) => line !== "").join(" ")
      return `${summary}\n\n${formatTodoTaskDetail(t)}`
    })
}

/**
 * Graph places a repeating task on the first date its pattern allows from the new
 * startDate, which is not always the date asked for: a weekly-on-Saturday task moved
 * to a Monday lands on the next Saturday. Say where it landed, rather than let the
 * caller assume the task is where they put it. The pattern is the usual reason but
 * not the only one (the returned date is in UTC), so the result reports the dates and
 * the rule without claiming which caused the gap.
 */
const describeReschedule = (
  requestedDue: string | undefined,
  updates: Record<string, unknown>,
  task: GraphTodoTask,
): string => {
  if (!requestedDue || updates.dueDateTime !== undefined || !task.recurrence) return ""
  const asked = toDateOnly(requestedDue)
  const landed = task.dueDateTime?.dateTime?.slice(0, 10)
  if (landed === undefined || landed === asked) return ""
  return `Graph put the task on ${landed}, not ${asked}. It repeats ${describeRecurrence(task.recurrence)}, and To Do moves a repeating task to the first date its pattern allows. Pass recurrence to change the pattern.`
}

/**
 * Completing a recurring task looks like a no-op in the response: Graph rolls the
 * same task id forward to the next due date and hands it back as notStarted, while
 * the completed occurrence becomes a separate task with a new id. Without a word of
 * explanation the caller sees "Task updated" over a notStarted task and reasonably
 * concludes the completion failed.
 */
const describeUpdate = (requestedStatus: string | undefined, task: GraphTodoTask): string => {
  if (requestedStatus !== "completed") return "Task updated."
  if (task.status === "completed") return "Task completed."
  if (!task.recurrence) return "Task updated."

  const next = Option(task.dueDateTime?.dateTime)
    .map((d) => ` Next occurrence is due ${d.slice(0, 10)}.`)
    .fold(
      () => "",
      (v) => v,
    )
  return `Occurrence completed, and the series rolled forward.${next} The completed occurrence is now a separate task; this id still refers to the live series.`
}

export const TASK_STATUSES = ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"] as const
export const TASK_IMPORTANCES = ["low", "normal", "high"] as const
export type TaskStatus = (typeof TASK_STATUSES)[number]
export type TaskImportance = (typeof TASK_IMPORTANCES)[number]

// Read the extension, apply the fields sent, and write the whole set back: a PATCH replaces every
// field (checked live 2026-10-30), so sending only the new ones would delete the rest. A missing
// extension is created instead.
const writeExtension = async (
  client: NonNullable<ReturnType<typeof requireClient>>,
  listId: string,
  taskId: string,
  extension: ExtensionInput,
): Promise<Either<string, Readonly<Record<string, ExtensionValue>>>> => {
  const existing = await client.getTodoTaskExtension(listId, taskId, extension.name)
  if (existing.isLeft()) {
    const error = existing.value as GraphApiError
    if (error.type !== "not_found") return Left(`reading the extension failed: ${error.message}`)
    const created = await client.createTodoTaskExtension(
      listId,
      taskId,
      toExtensionBody(extension.name, extension.fields),
    )
    return created.fold(
      (e) => Left(`creating the extension failed: ${e.message}`),
      () => Right(extension.fields),
    )
  }
  const merged = { ...extensionFields(existing.value as Record<string, unknown>), ...extension.fields }
  const tooBig = extensionFieldsProblem(merged)
  if (tooBig) return Left(`the extension can't take these fields: ${tooBig}`)
  const replaced = await client.replaceTodoTaskExtension(listId, taskId, extension.name, toExtensionPatch(merged))
  return replaced.fold(
    (e) => Left(`writing the extension failed: ${e.message}`),
    () => Right(merged),
  )
}

/**
 * The narrow update an unattended caller makes: status, importance and its own extension fields on
 * one task, nothing else. The same call twice leaves the task as one call does, which is why it
 * refuses to complete a repeating task: each completion rolls the series on by one more occurrence.
 *
 * Calls run in a fixed order. A read comes first only to check a completion; the task's PATCH comes
 * next; the extension last. A failure says what was saved and what wasn't, and the whole call can be
 * repeated safely.
 */
export const setTodoTaskState = async (params: {
  list_id: string
  task_id: string
  status?: TaskStatus
  importance?: TaskImportance
  extension?: ExtensionInput
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.status === undefined && params.importance === undefined && params.extension === undefined) {
    return Left(new UserError("Pass at least one of status, importance or extension."))
  }
  const extProblem = params.extension ? extensionProblem(params.extension) : undefined
  if (extProblem) return Left(new UserError(extProblem))

  if (params.status === "completed") {
    const existing = await client.getTodoTask(params.list_id, params.task_id)
    if (existing.isLeft()) {
      return Left(
        new UserError(`Failed to read the task, so nothing was changed: ${(existing.value as GraphApiError).message}`),
      )
    }
    if ((existing.value as GraphTodoTask).recurrence) {
      return Left(
        new UserError(
          "This task repeats, and completing it rolls the series on to its next occurrence each time, so " +
            "set_todo_task_state won't do it. Use update_todo_task, which is meant for attended use.",
        ),
      )
    }
  }

  const fields: Record<string, string> = {
    ...(params.status === undefined ? {} : { status: params.status }),
    ...(params.importance === undefined ? {} : { importance: params.importance }),
  }
  const changedTask = Object.keys(fields).length > 0
  if (changedTask) {
    const patched = await client.updateTodoTask(params.list_id, params.task_id, fields)
    if (patched.isLeft()) {
      return Left(
        new UserError(`Failed to update the task, so nothing was changed: ${(patched.value as GraphApiError).message}`),
      )
    }
  }

  const saved = changedTask ? `Set ${Object.keys(fields).join(" and ")}` : ""
  if (params.extension === undefined) return Right(`${saved} on task ${params.task_id}.`)

  const written = await writeExtension(client, params.list_id, params.task_id, params.extension)
  if (written.isLeft()) {
    const prefix = changedTask ? `${saved} on task ${params.task_id}, but ` : "Nothing was changed: "
    return Left(new UserError(`${prefix}${written.value as string}. Repeating the call is safe.`))
  }
  const now = extensionJson(written.value as Record<string, ExtensionValue>)
  return Right(
    `${changedTask ? `${saved}, and set` : "Set"} extension ${params.extension.name} on task ${params.task_id}. It now holds ${now}.`,
  )
}

type TodoClient = NonNullable<ReturnType<typeof requireClient>>

const RESTART =
  "Start again with delta: true, which returns every task and a new token. Graph sends this when a token " +
  "is unknown, damaged or too old."

// The three answers that mean a token can't be used again (seen or documented, 2026-10-30): any 410,
// the code syncStateNotFound whatever its status, and a 400 to a call whose only input is the token
// ("Badly formed token.").
const isRestart = (error: GraphApiError, withToken: boolean): boolean =>
  error.status === 410 ||
  error.graphErrorCode?.toLowerCase() === "syncstatenotfound" ||
  (withToken && error.status === 400)

type Fill = { readonly kind: "task"; readonly task: GraphTodoTask } | { readonly kind: "gone"; readonly id: string }

// One GET per task, run in turn. Delta records can carry only the changed properties, and a field left
// out can't be told from one cleared, so a line is printed from the task itself, never from a partial.
const fillTasks = async (
  client: TodoClient,
  listId: string,
  ids: ReadonlyArray<string>,
  extensionName: string | undefined,
): Promise<Either<string, ReadonlyArray<Fill>>> =>
  ids.reduce<Promise<Either<string, ReadonlyArray<Fill>>>>(
    async (acc, id) => {
      const soFar = await acc
      if (soFar.isLeft()) return soFar
      const got = await client.getTodoTask(listId, id, taskExpand(extensionName))
      if (got.isRight())
        return Right([...(soFar.value as ReadonlyArray<Fill>), { kind: "task", task: got.value as GraphTodoTask }])
      const error = got.value as GraphApiError
      return error.type === "not_found"
        ? Right([...(soFar.value as ReadonlyArray<Fill>), { kind: "gone", id }])
        : Left(`reading task ${id} failed: ${error.message}`)
    },
    Promise.resolve(Right([])),
  )

const hasCoreFields = (task: GraphTodoTask): boolean =>
  task.title !== undefined && task.status !== undefined && task.importance !== undefined

/**
 * A change feed over one list. With delta: true and no token, every task and a first token; with a
 * token, what was added, changed or removed since it, and the next token. The token is printed only
 * when every page and every task was read, so a run that fails part-way repeats the same window next
 * time instead of skipping it.
 */
const listTodoTaskChanges = async (
  client: TodoClient,
  params: { list_id: string; fetch_all_pages?: boolean; extension_name?: string; delta_token?: string },
): Promise<Either<UserError, string>> => {
  if (params.fetch_all_pages) {
    return Left(new UserError("A change feed always reads every page, so don't pass fetch_all_pages with delta."))
  }
  const withToken = params.delta_token !== undefined
  if (withToken && !isDeltaToken(params.delta_token!)) {
    return Left(new UserError(`That isn't a delta token this tool printed. ${RESTART}`))
  }

  const delta = await client.todoTaskDelta(params.list_id, params.delta_token)
  if (delta.isLeft()) {
    const error = delta.value as GraphApiError
    return Left(
      new UserError(
        isRestart(error, withToken)
          ? `Graph won't accept this delta token (${error.message}). ${RESTART}`
          : `Failed to read changes, so no token was printed: ${error.message}`,
      ),
    )
  }
  const { records, deltaToken } = delta.value as TodoDeltaResult

  // Graph doesn't promise an order, and one task can appear more than once; such a task is read afresh.
  const byId = records.reduce(
    (m, r) => m.set(r.id, [...(m.get(r.id) ?? []), r]),
    new Map<string, ReadonlyArray<TodoDeltaRecord>>(),
  )
  const entries = [...byId.entries()]
  const removedOnly = entries.filter(([, rs]) => rs.every((r) => r["@removed"] !== undefined))
  const live = entries.filter(([, rs]) => !rs.every((r) => r["@removed"] !== undefined))
  // With a token, every live task is read afresh. On the first call, the delta record is printed as it
  // came, unless it is missing a core field or appeared twice. Delta can't return extensions (live,
  // 2026-10-30), so the first call prints none, and the GETs it makes leave the extension out too.
  const toRead = live.filter(([, rs]) => withToken || rs.length > 1 || !hasCoreFields(rs[0]!)).map(([id]) => id)
  const asIs = live.filter(([id]) => !toRead.includes(id)).map(([, rs]) => rs[0] as GraphTodoTask)
  const extensionName = withToken ? params.extension_name : undefined

  const filled = await fillTasks(client, params.list_id, toRead, extensionName)
  if (filled.isLeft()) {
    return Left(new UserError(`Failed to read changes, so no token was printed: ${filled.value as string}`))
  }
  const fills = filled.value as ReadonlyArray<Fill>
  const tasks = [...asIs, ...fills.flatMap((f) => (f.kind === "task" ? [f.task] : []))]
  const removed = [
    ...removedOnly.map(([id, rs]) => ({ id, reason: rs.find((r) => r["@removed"]?.reason)?.["@removed"]?.reason })),
    ...fills.flatMap((f) => (f.kind === "gone" ? [{ id: f.id }] : [])),
  ]
  const notes =
    !withToken && params.extension_name !== undefined
      ? [
          "the first delta call can't return extensions, so these lines carry no Ext. Read the list once with " +
            "list_todo_tasks and extension_name to get them; calls with a token include them.",
        ]
      : []
  return Right(formatTodoTaskChanges({ first: !withToken, tasks, removed, deltaToken, extensionName, notes }))
}
