import type { GraphApiError, GraphRequest } from "@sapientsai/ms-graph-core"
import type { Either } from "functype/either"
import { z } from "zod"

// Generic Microsoft Graph passthrough + $batch, ported from microsoft-mcp-server onto
// core's createGraphRequest. Tool names match the gateway's for drop-in compatibility.
// (The gateway's Azure-Management apiType is omitted: the app-only token is Graph-scoped;
// ARM passthrough would need a separately-scoped token — deferred follow-up.)

const METHODS = ["GET", "POST", "PUT", "PATCH", "DELETE"] as const
const API_VERSIONS = ["v1.0", "beta"] as const

const unwrap = <T>(result: Either<GraphApiError, T>): string =>
  result.fold(
    (error) => {
      throw new Error(error.message)
    },
    (data) => JSON.stringify(data, null, 2),
  )

const withQuery = (path: string, queryParams?: Record<string, string>): string => {
  if (!queryParams || Object.keys(queryParams).length === 0) return path
  const qs = new URLSearchParams(queryParams).toString()
  return path.includes("?") ? `${path}&${qs}` : `${path}?${qs}`
}

export const buildMicrosoftGraphTool = (graph: GraphRequest, extraInstructions?: string) => ({
  name: "microsoft_graph",
  description: `Execute a Microsoft Graph API request. Access Microsoft 365 data — users, mail, calendar, files, sites, and more.${
    extraInstructions ? ` ${extraInstructions}` : ""
  }`,
  parameters: z.object({
    path: z.string().describe("API endpoint path, e.g. /me, /users, /me/messages"),
    method: z.enum(METHODS).default("GET").describe("HTTP method"),
    api_version: z.enum(API_VERSIONS).default("v1.0").describe("Graph API version"),
    query_params: z
      .record(z.string(), z.string())
      .optional()
      .describe("OData query parameters ($select, $filter, $top, $orderby, etc.)"),
    body: z
      .union([z.record(z.string(), z.unknown()), z.string()])
      .optional()
      .describe("Request body for POST/PUT/PATCH"),
  }),
  execute: async (args: {
    path: string
    method: (typeof METHODS)[number]
    api_version: (typeof API_VERSIONS)[number]
    query_params?: Record<string, string>
    body?: Record<string, unknown> | string
  }): Promise<string> => {
    const result = await graph.request<unknown>(args.method, withQuery(args.path, args.query_params), {
      version: args.api_version,
      body: args.body,
    })
    return unwrap(result)
  },
})

export const buildMicrosoftGraphBatchTool = (graph: GraphRequest) => ({
  name: "microsoft_graph_batch",
  description:
    "Execute multiple Microsoft Graph API requests in a single batch call (max 20). Bulk operations like " +
    "creating folder trees or many reads. Individual request failures don't fail the batch — each response " +
    "carries its own status code.",
  parameters: z.object({
    requests: z
      .array(
        z.object({
          id: z.string().describe("Unique ID to correlate with the response"),
          method: z.enum(METHODS).describe("HTTP method"),
          url: z.string().describe("Relative API path, e.g. /me/drive/root/children"),
          headers: z
            .record(z.string(), z.string())
            .optional()
            .describe("Request headers (Content-Type auto-added for bodies)"),
          body: z
            .union([z.record(z.string(), z.unknown()), z.string()])
            .optional()
            .describe("Request body for POST/PUT/PATCH"),
          dependsOn: z.array(z.string()).optional().describe("IDs of requests that must complete first"),
        }),
      )
      .min(1)
      .max(20)
      .describe("Array of requests (max 20)"),
    api_version: z.enum(API_VERSIONS).default("v1.0").describe("Graph API version"),
  }),
  execute: async (args: {
    requests: ReadonlyArray<{
      id: string
      method: (typeof METHODS)[number]
      url: string
      headers?: Record<string, string>
      body?: Record<string, unknown> | string
      dependsOn?: string[]
    }>
    api_version: (typeof API_VERSIONS)[number]
  }): Promise<string> => {
    const batchRequests = args.requests.map((req) => {
      const normalized: Record<string, unknown> = { id: req.id, method: req.method, url: req.url }
      if (req.body !== undefined) {
        normalized.body = typeof req.body === "string" ? JSON.parse(req.body) : req.body
        const headers: Record<string, string> = req.headers ? { ...req.headers } : {}
        if (!Object.keys(headers).some((k) => k.toLowerCase() === "content-type")) {
          headers["Content-Type"] = "application/json"
        }
        normalized.headers = headers
      } else if (req.headers) {
        normalized.headers = req.headers
      }
      if (req.dependsOn && req.dependsOn.length > 0) normalized.dependsOn = req.dependsOn
      return normalized
    })

    const result = await graph.request<unknown>("POST", "/$batch", {
      version: args.api_version,
      body: { requests: batchRequests },
    })
    return unwrap(result)
  },
})
