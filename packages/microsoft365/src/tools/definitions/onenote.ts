// OneNote tool definitions.

import { z } from "zod"

import {
  copyOnenotePage,
  createOnenoteNotebook,
  createOnenotePage,
  createOnenoteSection,
  deleteOnenotePage,
  getOnenotePageContent,
  listOnenoteNotebooks,
  listOnenotePages,
  listOnenoteSections,
  updateOnenotePageContent,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const onenoteTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_onenote_notebooks",
    description: "List OneNote notebooks",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenoteNotebooks(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_onenote_sections",
    description: "List sections in a OneNote notebook",
    parameters: z.object({
      notebook_id: z.string().describe("Notebook ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenoteSections(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_onenote_pages",
    description: "List pages in a OneNote section",
    parameters: z.object({
      section_id: z.string().describe("Section ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listOnenotePages(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_onenote_page_content",
    description: "Get OneNote page content as HTML",
    parameters: z.object({
      page_id: z.string().describe("Page ID"),
    }),
    execute: async (params) => unwrapResult(await getOnenotePageContent(params)),
    domain: "onenote",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_onenote_page",
    description:
      "Create a OneNote page in a section. Content is HTML — OneNote supports a constrained subset and silently drops unsupported CSS/tags.",
    parameters: z.object({
      section_id: z.string().describe("Section ID to create the page in"),
      title: z.string().describe("Page title"),
      content: z.string().describe("Page body as HTML (e.g. <p>Hello</p>)"),
    }),
    execute: async (params) => unwrapResult(await createOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "update_onenote_page_content",
    description:
      "Append to or modify a OneNote page's content without rewriting it. Targets an element (default 'body') and applies an action. Content is HTML.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to update"),
      content: z.string().describe("HTML content to apply"),
      action: z.enum(["append", "insert", "prepend", "replace"]).optional().describe("Update action (default: append)"),
      target: z.string().optional().describe("Target element data-id or known name (default: body)"),
      position: z.enum(["before", "after"]).optional().describe("Position relative to target, for insert"),
    }),
    execute: async (params) => unwrapResult(await updateOnenotePageContent(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "create_onenote_section",
    description: "Create a new section in a OneNote notebook",
    parameters: z.object({
      notebook_id: z.string().describe("Notebook ID to create the section in"),
      display_name: z.string().describe("Section name"),
    }),
    execute: async (params) => unwrapResult(await createOnenoteSection(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "create_onenote_notebook",
    description: "Create a new OneNote notebook",
    parameters: z.object({
      display_name: z.string().describe("Notebook name"),
    }),
    execute: async (params) => unwrapResult(await createOnenoteNotebook(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "copy_onenote_page",
    description: "Copy a OneNote page to another section. Runs asynchronously in OneNote.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to copy"),
      section_id: z.string().describe("Destination section ID"),
    }),
    execute: async (params) => unwrapResult(await copyOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
  },
  {
    name: "delete_onenote_page",
    description: "Delete a OneNote page. This cannot be undone via the API.",
    parameters: z.object({
      page_id: z.string().describe("Page ID to delete"),
    }),
    execute: async (params) => unwrapResult(await deleteOnenotePage(params)),
    domain: "onenote",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
