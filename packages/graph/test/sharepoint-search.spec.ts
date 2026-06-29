import type { GraphRequest } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { resolveSharePointSearchConfig } from "../src/config"
import { buildSharePointSearchTool, type SharePointSearchConfig } from "../src/tools/sharepoint-search"

const hit = (over: Record<string, unknown> = {}) => ({
  resource: {
    id: "item-1",
    name: "Report.docx",
    webUrl: "https://sp/Report.docx",
    lastModifiedDateTime: "2026-01-01T00:00:00Z",
    size: 1234,
    file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
    parentReference: { driveId: "drive-1", siteId: "site-1" },
    ...over,
  },
  _summary: "matched <c0>keyword</c0>",
})

const searchResult = (hits: unknown[]) => Right({ value: [{ hitsContainers: [{ hits }] }] })

const fakeGraph = (impl?: GraphRequest["request"]) => {
  const request = vi.fn(impl ?? (() => Promise.resolve(searchResult([hit()]))))
  return { graph: { request, requestPaginated: vi.fn() } as unknown as GraphRequest, request }
}

const baseConfig: SharePointSearchConfig = { region: "NAM" }

describe("resolveSharePointSearchConfig", () => {
  it("defaults region to NAM and reads SITE_ID/SITE_URL", () => {
    expect(resolveSharePointSearchConfig({})).toEqual({
      region: "NAM",
      defaultSiteId: undefined,
      defaultSiteUrl: undefined,
    })
    const cfg = resolveSharePointSearchConfig({
      GRAPH_SEARCH_REGION: "EUR",
      SITE_ID: "s1",
      SITE_URL: "https://contoso.sharepoint.com/sites/x/",
    })
    expect(cfg).toEqual({
      region: "EUR",
      defaultSiteId: "s1",
      defaultSiteUrl: "https://contoso.sharepoint.com/sites/x",
    })
  })
})

describe("sharepoint_search tool", () => {
  it("POSTs a driveItem search to /search/query with the region", async () => {
    const { graph, request } = fakeGraph()
    await buildSharePointSearchTool(graph, baseConfig).execute({ query: "budget", top: 10 })

    const [method, path, opts] = request.mock.calls[0] as [
      string,
      string,
      { body: { requests: Array<Record<string, unknown>> } },
    ]
    expect(method).toBe("POST")
    expect(path).toBe("/search/query")
    const req = opts.body.requests[0]
    expect(req).toMatchObject({ entityTypes: ["driveItem"], region: "NAM", size: 10 })
    expect((req.query as { queryString: string }).queryString).toBe("budget")
  })

  it("adds a filetype KQL clause and maps hits to results", async () => {
    const { graph, request } = fakeGraph()
    const out = await buildSharePointSearchTool(graph, baseConfig).execute({
      query: "q",
      top: 5,
      fileTypes: ["docx", ".pdf"],
    })
    const qs = (
      request.mock.calls[0] as [string, string, { body: { requests: Array<{ query: { queryString: string } }> } }]
    )[2].body.requests[0].query.queryString
    expect(qs).toBe("q (filetype:docx OR filetype:pdf)")

    const parsed = JSON.parse(out) as {
      results: Array<{ name: string; driveId: string; driveItemId: string }>
      totalCount: number
    }
    expect(parsed.totalCount).toBe(1)
    expect(parsed.results[0]).toMatchObject({ name: "Report.docx", driveId: "drive-1", driveItemId: "item-1" })
  })

  it("uses the default site URL as a KQL site: clause without an API call", async () => {
    const { graph, request } = fakeGraph()
    await buildSharePointSearchTool(graph, {
      region: "NAM",
      defaultSiteUrl: "https://contoso.sharepoint.com/sites/x",
    }).execute({
      query: "q",
      top: 10,
    })
    // exactly one call — the search POST; no /sites/{id} resolution
    expect(request).toHaveBeenCalledOnce()
    const qs = (
      request.mock.calls[0] as [string, string, { body: { requests: Array<{ query: { queryString: string } }> } }]
    )[2].body.requests[0].query.queryString
    expect(qs).toBe("q site:https://contoso.sharepoint.com/sites/x")
  })

  it("resolves an explicit siteId to its webUrl for the site: clause", async () => {
    const request = vi.fn((method: string, path: string) =>
      path.startsWith("/sites/")
        ? Promise.resolve(Right({ webUrl: "https://contoso.sharepoint.com/sites/y" }))
        : Promise.resolve(searchResult([])),
    )
    const graph = { request, requestPaginated: vi.fn() } as unknown as GraphRequest
    await buildSharePointSearchTool(graph, { region: "NAM" }).execute({ query: "q", siteId: "site-y", top: 10 })

    expect((request.mock.calls[0] as [string, string])[1]).toBe("/sites/site-y")
    const qs = (
      request.mock.calls[1] as [string, string, { body: { requests: Array<{ query: { queryString: string } }> } }]
    )[2].body.requests[0].query.queryString
    expect(qs).toContain("site:https://contoso.sharepoint.com/sites/y")
  })

  it("throws when the search request fails", async () => {
    const { graph } = fakeGraph(() => Promise.resolve(Left({ type: "forbidden", message: "no access" })))
    await expect(buildSharePointSearchTool(graph, baseConfig).execute({ query: "q", top: 10 })).rejects.toThrow(
      "no access",
    )
  })
})
