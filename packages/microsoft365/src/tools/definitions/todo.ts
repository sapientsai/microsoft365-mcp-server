// To Do tool definitions.

import { z } from "zod"

import { createTodoTask, listTodoLists, listTodoTasks, updateTodoTask } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

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
    description: "Create a new To Do task",
    parameters: z.object({
      list_id: z.string().describe("To Do list ID"),
      title: z.string().describe("Task title"),
      body: z.string().optional().describe("Task body/notes"),
      due_date: z.string().optional().describe("Due date (ISO format)"),
      importance: z.string().optional().describe("Importance: low, normal, or high"),
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
    }),
    execute: async (params) => unwrapResult(await updateTodoTask(params)),
    domain: "todo",
    readOnly: false,
  },
]
