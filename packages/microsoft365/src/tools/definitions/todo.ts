// To Do tool definitions.

import { z } from "zod"

import { DAYS_OF_WEEK, RECURRENCE_PATTERN_TYPES, RECURRENCE_RANGE_TYPES, WEEK_INDEXES } from "../../utils/recurrence"
import { createTodoTask, deleteTodoTask, listTodoLists, listTodoTasks, updateTodoTask } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

// The flat shape here mirrors Graph's patternedRecurrence without making the caller
// nest two objects. Which fields are required depends on the pattern — the builder
// enforces that and names what is missing, so the description stays short.
const RECURRENCE_PARAM = z
  .object({
    pattern: z.enum(RECURRENCE_PATTERN_TYPES).describe("How the task repeats"),
    interval: z
      .number()
      .int()
      .positive()
      .optional()
      .describe("Units between occurrences, in the pattern's unit. Default 1. Quarterly is absoluteMonthly with 3."),
    days_of_week: z
      .array(z.enum(DAYS_OF_WEEK))
      .optional()
      .describe("Required for weekly, relativeMonthly and relativeYearly"),
    day_of_month: z
      .number()
      .int()
      .min(1)
      .max(31)
      .optional()
      .describe("Required for absoluteMonthly and absoluteYearly"),
    month: z.number().int().min(1).max(12).optional().describe("Required for absoluteYearly and relativeYearly"),
    index: z
      .enum(WEEK_INDEXES)
      .optional()
      .describe("Which week of the month, for the relative patterns. Default first"),
    first_day_of_week: z.enum(DAYS_OF_WEEK).optional().describe("For weekly patterns. Default monday"),
    range_type: z.enum(RECURRENCE_RANGE_TYPES).optional().describe("noEnd (default), endDate, or numbered"),
    start_date: z.string().optional().describe("YYYY-MM-DD. Defaults to the task's due date"),
    end_date: z.string().optional().describe("YYYY-MM-DD. Required when range_type is endDate"),
    number_of_occurrences: z.number().int().positive().optional().describe("Required when range_type is numbered"),
    recurrence_time_zone: z.string().optional().describe("Time zone for the range, e.g. 'W. Australia Standard Time'"),
  })
  .optional()

export const todoTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_todo_lists",
    description: "List Microsoft To Do task lists",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTodoLists(params)),
    domain: "todo",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_todo_tasks",
    description: "List tasks in a To Do list",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTodoTasks(params)),
    domain: "todo",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_todo_task",
    description: "Create a new To Do task, optionally recurring",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      title: z.string().describe("Task title"),
      body: z.string().optional().describe("Task body/notes"),
      due_date: z.string().optional().describe("Due date (ISO format). Required when recurrence is set"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
      recurrence: RECURRENCE_PARAM.describe("Repeat pattern. Omit for a one-off task"),
    }),
    execute: async (params) => unwrapResult(await createTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },
  {
    name: "update_todo_task",
    description: "Update a To Do task",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
      title: z.string().optional().describe("New title"),
      status: z.string().optional().describe("Status: notStarted, inProgress, completed, waitingOnOthers, deferred"),
      due_date: z.string().optional().describe("New due date (ISO format)"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
      body: z.string().optional().describe("New body/notes"),
      recurrence: RECURRENCE_PARAM.describe("Replace the repeat pattern"),
      clear_recurrence: z.boolean().optional().describe("Remove the repeat pattern, making the task one-off"),
    }),
    execute: async (params) => unwrapResult(await updateTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },
  {
    name: "delete_todo_task",
    description: "Delete a To Do task permanently. To Do has no recycle bin — this cannot be undone",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      task_id: z.string().describe("Task ID"),
      force: z
        .boolean()
        .optional()
        .describe(
          "Required to delete a repeating task, which ends the whole series rather than one occurrence. " +
            "To stop it repeating but keep the task, use clear_recurrence on update_todo_task instead.",
        ),
    }),
    execute: async (params) => unwrapResult(await deleteTodoTask(params)),
    domain: "todo",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
