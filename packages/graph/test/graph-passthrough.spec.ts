import type { GraphRequest } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { buildMicrosoftGraphBatchTool, buildMicrosoftGraphTool } from "../src/tools/graph-passthrough"

const fakeGraph = (impl?: GraphRequest["request"]): { graph: GraphRequest; request: ReturnType<typeof vi.fn> } => {
  const request = vi.fn(impl ?? (() => Promise.resolve(Right({ ok: true }))))
  return { graph: { request, requestPaginated: vi.fn() } as unknown as GraphRequest, request }
}

describe("microsoft_graph passthrough", () => {
  it("forwards method/version/body and appends query params to the path", async () => {
    const { graph, request } = fakeGraph()
    const tool = buildMicrosoftGraphTool(graph)

    const out = await tool.execute({
      path: "/me/messages",
      method: "GET",
      api_version: "beta",
      query_params: { $select: "subject", $top: "5" },
    })

    expect(request).toHaveBeenCalledWith("GET", "/me/messages?%24select=subject&%24top=5", {
      version: "beta",
      body: undefined,
    })
    expect(JSON.parse(out)).toEqual({ ok: true })
  })

  it("merges query params onto a path that already has a query string", async () => {
    const { graph, request } = fakeGraph()
    await buildMicrosoftGraphTool(graph).execute({
      path: "/me/calendarView?startDateTime=x",
      method: "GET",
      api_version: "v1.0",
      query_params: { $top: "1" },
    })
    expect((request.mock.calls[0] as [string, string])[1]).toBe("/me/calendarView?startDateTime=x&%24top=1")
  })

  it("appends custom instructions to the description", () => {
    expect(buildMicrosoftGraphTool(fakeGraph().graph, "Prefer read-only.").description).toContain("Prefer read-only.")
  })

  it("throws with the Graph error message on a Left result", async () => {
    const { graph } = fakeGraph(() => Promise.resolve(Left({ type: "forbidden", message: "Access denied" })))
    await expect(
      buildMicrosoftGraphTool(graph).execute({ path: "/users", method: "GET", api_version: "v1.0" }),
    ).rejects.toThrow("Access denied")
  })
})

describe("microsoft_graph_batch", () => {
  it("POSTs to /$batch and auto-adds Content-Type for requests with a body", async () => {
    const { graph, request } = fakeGraph()
    const tool = buildMicrosoftGraphBatchTool(graph)

    await tool.execute({
      api_version: "v1.0",
      requests: [
        { id: "1", method: "GET", url: "/me" },
        { id: "2", method: "POST", url: "/me/messages", body: { subject: "hi" }, dependsOn: ["1"] },
      ],
    })

    const [method, path, opts] = request.mock.calls[0] as [string, string, { body: { requests: unknown[] } }]
    expect(method).toBe("POST")
    expect(path).toBe("/$batch")
    const reqs = opts.body.requests as Array<Record<string, unknown>>
    expect(reqs[1]).toMatchObject({
      id: "2",
      headers: { "Content-Type": "application/json" },
      body: { subject: "hi" },
      dependsOn: ["1"],
    })
    // GET request gets no auto Content-Type
    expect(reqs[0]).not.toHaveProperty("headers")
  })

  it("parses a string body into JSON for the batch payload", async () => {
    const { graph, request } = fakeGraph()
    await buildMicrosoftGraphBatchTool(graph).execute({
      api_version: "v1.0",
      requests: [{ id: "1", method: "POST", url: "/x", body: '{"a":1}' }],
    })
    const reqs = (request.mock.calls[0] as [string, string, { body: { requests: Array<{ body: unknown }> } }])[2].body
      .requests
    expect(reqs[0].body).toEqual({ a: 1 })
  })
})
