// Planner tool definitions.

import { z } from "zod"

import {
  createPlannerBucket,
  createPlannerTask,
  getPlannerTask,
  listPlannerBuckets,
  listPlannerTasks,
  listPlans,
  updatePlannerTask,
  updatePlannerTaskDetails,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const plannerTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_plans",
    description: "List all Planner plans visible to you, aggregated across your group memberships",
    parameters: z.object({}),
    execute: async () => unwrapResult(await listPlans()),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_planner_tasks",
    description: "List tasks in a Planner plan",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listPlannerTasks(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_planner_task",
    description: "Get detailed Planner task information",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
    }),
    execute: async (params) => unwrapResult(await getPlannerTask(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_planner_buckets",
    description: "List the buckets (columns) in a Planner plan. Use a bucket ID with create_planner_task.",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
    }),
    execute: async (params) => unwrapResult(await listPlannerBuckets(params)),
    domain: "planner",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_planner_bucket",
    description: "Create a new bucket (column) in a Planner plan",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      name: z.string().describe("Bucket name"),
    }),
    execute: async (params) => unwrapResult(await createPlannerBucket(params)),
    domain: "planner",
    readOnly: false,
  },
  {
    name: "create_planner_task",
    description: "Create a new Planner task",
    parameters: z.object({
      plan_id: z.string().describe("Plan ID"),
      title: z.string().describe("Task title"),
      bucket_id: z.string().optional().describe("Bucket ID"),
      due_date: z.string().optional().describe("Due date (ISO format)"),
      assignments: z.string().optional().describe("Comma-separated user IDs to assign"),
    }),
    execute: async (params) => unwrapResult(await createPlannerTask(params)),
    domain: "planner",
    readOnly: false,
  },
  {
    name: "update_planner_task",
    description:
      "Update a Planner task (title, percent_complete — 100 closes it — due date, priority). The ETag " +
      "is auto-fetched if omitted; pass one only to enforce optimistic concurrency.",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
      etag: z.string().optional().describe("Task ETag — omit to auto-fetch; pass to enforce concurrency"),
      title: z.string().optional().describe("New title"),
      percent_complete: z.number().optional().describe("Completion percentage (0-100; 100 marks complete)"),
      due_date: z.string().optional().describe("New due date (ISO format)"),
      priority: z.number().optional().describe("Priority (0-10)"),
    }),
    execute: async (params) => unwrapResult(await updatePlannerTask(params)),
    domain: "planner",
    readOnly: false,
  },
  {
    name: "update_planner_task_details",
    description:
      "Update a Planner task's details: description, checklist items, and references (links). The ETag " +
      "is fetched automatically. Add, update, or remove entries — updates merge (omitted fields keep " +
      "their current value); update/remove of a missing item is reported as skipped, not a silent no-op. " +
      "Checklist items are addressed by their GUID (from get_planner_task); references by their URL.",
    parameters: z.object({
      task_id: z.string().describe("Task ID"),
      description: z.string().optional().describe("Task description / notes"),
      preview_type: z
        .enum(["automatic", "noPreview", "checklist", "description", "reference"])
        .optional()
        .describe("What shows on the task card face"),
      add_checklist: z
        .array(z.object({ title: z.string(), isChecked: z.boolean().optional() }))
        .optional()
        .describe("Checklist items to add"),
      update_checklist: z
        .array(z.object({ id: z.string(), title: z.string().optional(), isChecked: z.boolean().optional() }))
        .optional()
        .describe("Checklist items to edit, addressed by GUID (e.g. toggle isChecked)"),
      remove_checklist: z.array(z.string()).optional().describe("Checklist item GUIDs to remove"),
      add_references: z
        .array(z.object({ url: z.string(), alias: z.string().optional() }))
        .optional()
        .describe("Reference links to add"),
      update_references: z
        .array(z.object({ url: z.string(), alias: z.string().optional() }))
        .optional()
        .describe("Reference links to edit, addressed by URL (e.g. rename the alias)"),
      remove_references: z.array(z.string()).optional().describe("Reference URLs to remove"),
    }),
    execute: async (params) => unwrapResult(await updatePlannerTaskDetails(params)),
    domain: "planner",
    readOnly: false,
  },
]
