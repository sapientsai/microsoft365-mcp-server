// Graph query (escape hatch) tool definitions.

import { z } from "zod"

import { graphQuery } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { unwrapResult } from "./shared"

export const queryTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "graph_query",
    description: "Execute an arbitrary Microsoft Graph API query. Use this for operations not covered by other tools.",
    parameters: z.object({
      method: z.string().describe("HTTP method: GET, POST, PUT, PATCH, or DELETE"),
      path: z.string().describe("Graph API path (e.g., /me/memberOf)"),
      body: z.string().optional().describe("JSON request body as a string"),
      version: z.string().optional().describe("API version: v1.0 or beta (default: v1.0)"),
      headers: z
        .record(z.string(), z.string())
        .optional()
        .describe('Extra request headers, e.g. { "If-Match": "<etag>" } for concurrency-controlled writes'),
    }),
    execute: async (params) => unwrapResult(await graphQuery(params)),
    domain: "query",
    readOnly: false,
    annotations: { destructiveHint: true, openWorldHint: true },
  },
]
