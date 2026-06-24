import { describe, expect, it, vi } from "vitest"

import { resolveAiSearchConfig } from "../src/config"
import { aiSearchFetch, parseAiSearchError } from "../src/search/ai-search-client"
import { buildAiSearchTool } from "../src/tools/ai-search"

const config = {
  endpoint: "https://x.search.windows.net",
  apiKey: "KEY",
  indexName: "idx",
  semanticConfiguration: "sem",
  vectorFields: "vec",
}

const searchResponse = (body: unknown, ok = true, status = 200) =>
  ({
    ok,
    status,
    statusText: ok ? "OK" : "Error",
    json: () => Promise.resolve(body),
  }) as Response

describe("resolveAiSearchConfig", () => {
  it("returns undefined unless endpoint + api key + index are all set", () => {
    expect(resolveAiSearchConfig({})).toBeUndefined()
    expect(resolveAiSearchConfig({ AZURE_AI_SEARCH_ENDPOINT: "e", AZURE_AI_SEARCH_API_KEY: "k" })).toBeUndefined()
  })

  it("resolves a full config and strips a trailing slash from the endpoint", () => {
    const cfg = resolveAiSearchConfig({
      AZURE_AI_SEARCH_ENDPOINT: "https://x.search.windows.net/",
      AZURE_AI_SEARCH_API_KEY: "k",
      AZURE_AI_SEARCH_INDEX: "idx",
      AZURE_AI_SEARCH_SEMANTIC_CONFIG: "sem",
    })
    expect(cfg?.endpoint).toBe("https://x.search.windows.net")
    expect(cfg?.semanticConfiguration).toBe("sem")
    expect(cfg?.vectorFields).toBeUndefined()
  })
})

describe("aiSearchFetch / parseAiSearchError", () => {
  it("sends the api-key header and returns the response on success", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(searchResponse({ value: [] })))
    const result = await aiSearchFetch("https://x/search", "KEY", {}, fetchImpl as unknown as typeof fetch)
    expect(result.isRight()).toBe(true)
    expect((fetchImpl.mock.calls[0] as [string, RequestInit])[1].headers).toMatchObject({ "api-key": "KEY" })
  })

  it("maps a non-OK response to an api error with status", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve(searchResponse({ error: { code: "Forbidden", message: "bad key" } }, false, 403)),
    )
    const result = await aiSearchFetch("https://x/search", "KEY", {}, fetchImpl as unknown as typeof fetch)
    expect(result.isLeft()).toBe(true)
    expect(result.value as { type: string; status: number }).toMatchObject({ type: "api", status: 403 })
  })

  it("parseAiSearchError prefixes the code when present", async () => {
    const msg = await parseAiSearchError(searchResponse({ error: { code: "X", message: "boom" } }, false, 400))
    expect(msg).toBe("X: boom")
  })
})

describe("azure_ai_search tool", () => {
  it("builds a semantic search body with captions/answers + vector query", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(searchResponse({ value: [], "@odata.count": 0 })))
    const tool = buildAiSearchTool(config, fetchImpl as unknown as typeof fetch)

    await tool.execute({
      query: "hello",
      queryType: "semantic",
      vectorSearch: true,
      top: 5,
      includeTotalCount: true,
    })

    const [url, init] = fetchImpl.mock.calls[0] as [string, RequestInit]
    expect(url).toContain("/indexes/idx/docs/search?api-version=2025-09-01")
    const body = JSON.parse(String(init.body)) as Record<string, unknown>
    expect(body).toMatchObject({
      search: "hello",
      queryType: "semantic",
      semanticConfiguration: "sem",
      captions: "extractive",
    })
    expect((body.vectorQueries as unknown[]).length).toBe(1)
  })

  it("maps hits to scored results with captions/highlights", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve(
        searchResponse({
          value: [
            {
              "@search.score": 1.5,
              "@search.captions": [{ text: "a caption" }],
              "@search.highlights": { title: ["<em>hi</em>"] },
              title: "Doc",
            },
          ],
        }),
      ),
    )
    const out = await buildAiSearchTool(config, fetchImpl as unknown as typeof fetch).execute({
      query: "q",
      queryType: "simple",
      vectorSearch: false,
      top: 10,
      includeTotalCount: false,
    })
    const parsed = JSON.parse(out) as {
      results: Array<{ score: number; document: { title: string }; captions: string[] }>
    }
    expect(parsed.results[0]).toMatchObject({ score: 1.5, document: { title: "Doc" }, captions: ["a caption"] })
  })

  it("rejects semantic queries when no semantic configuration is set", async () => {
    const tool = buildAiSearchTool({ ...config, semanticConfiguration: undefined }, vi.fn() as unknown as typeof fetch)
    await expect(
      tool.execute({ query: "q", queryType: "semantic", vectorSearch: false, top: 10, includeTotalCount: false }),
    ).rejects.toThrow("semantic configuration")
  })

  it("throws the search error message on a failed request", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(searchResponse({ error: { message: "index missing" } }, false, 404)))
    await expect(
      buildAiSearchTool(config, fetchImpl as unknown as typeof fetch).execute({
        query: "q",
        queryType: "simple",
        vectorSearch: false,
        top: 10,
        includeTotalCount: false,
      }),
    ).rejects.toThrow("index missing")
  })
})
