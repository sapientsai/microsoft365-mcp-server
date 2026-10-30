// Parameter schemas for the To Do tools that need a test of their own. index.ts can't be loaded in
// a test (loading it starts the server), so these live here and index.ts imports them.

import { z } from "zod"

import { TASK_IMPORTANCES, TASK_STATUSES } from "./todo-tools"

// A caller's own data on a task, as a Graph open extension. Use one fixed name per app: Graph warns
// against naming extensions in a variable manner.
export const TODO_EXTENSION_PARAM = z
  .object({
    name: z.string().describe("Reverse-DNS extension name, one fixed name per app, e.g. com.example.assistant"),
    fields: z
      .record(z.string(), z.union([z.string(), z.number(), z.boolean()]))
      .describe("At most 20 flat fields (string, number or boolean), at most 2 KB as JSON"),
  })
  .optional()

// Strict, unlike the other tools' schemas: zod drops unknown keys by default, so a call carrying a
// title or categories would set the status and report success with the rest silently ignored. This
// tool promises to change nothing but these fields, so anything else is refused.
export const SET_TODO_TASK_STATE_PARAMS = z
  .object({
    list_id: z.string().describe("To Do list ID"),
    task_id: z.string().describe("Task ID"),
    status: z.enum(TASK_STATUSES).optional().describe("New status"),
    importance: z.enum(TASK_IMPORTANCES).optional().describe("New importance"),
    extension: TODO_EXTENSION_PARAM.describe(
      "Fields to set in your extension. Fields you leave out keep their values; a missing extension is created",
    ),
  })
  .strict()
