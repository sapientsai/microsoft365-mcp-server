import { getSharedBreaker, resetSharedBreakerForTests } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import { setAccessTokenTool } from "../src/tools/auth-tools"
import { uploadFile, uploadFileFromPath } from "../src/tools/files-tools"
import { GRAPH_QUERY_HEADERS, graphQuery } from "../src/tools/graph-query-tools"
import { updatePlannerTask } from "../src/tools/planner-tools"

// A value fetch can't put in a header used to throw inside the resilience layer, where it counted
// as a network failure: a couple of bad calls could open the shared breaker and fail every tool.
// Each tool now refuses such a value itself, before any request.

const json = (value: unknown, status = 200) =>
  new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } })

const fetchSpy = vi.fn((_url: string, _init?: RequestInit) => Promise.resolve(json({ ok: true })))
const message = (result: { value: unknown }): string => (result.value as Error).message

beforeEach(() => {
  fetchSpy.mockClear()
  vi.stubGlobal("fetch", fetchSpy)
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  resetSharedBreakerForTests()
  initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("TOKEN")) })
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  expect(getSharedBreaker().getState().failures).toBe(0)
})

describe("graph_query method", () => {
  it.each(["CONNECT", "TRACE", "OPTIONS", "HEAD", "PROPFIND", "", "GE T"])("refuses %j", async (method) => {
    const result = await graphQuery({ method, path: "/me" })
    expect(message(result)).toContain("method must be one of GET, POST, PUT, PATCH, DELETE")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("sends a lower-case post as POST, with its body", async () => {
    const result = await graphQuery({ method: " post ", path: "/me/events", body: '{"subject":"x"}' })
    expect(result.isRight()).toBe(true)
    const init = fetchSpy.mock.calls[0]?.[1]
    expect(init?.method).toBe("POST")
    expect(init?.body).toBe('{"subject":"x"}')
  })

  it("applies the draft rule to a lower-case method", async () => {
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    const result = await graphQuery({ method: "post", path: "/me/sendMail", body: "{}" })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})

describe("graph_query headers", () => {
  it.each(["Authorization", "authorization", "Host", "X-HTTP-Method-Override", "Cookie", "Upgrade"])(
    "refuses %s, which isn't on the list",
    async (name) => {
      const result = await graphQuery({ method: "GET", path: "/me", headers: { [name]: "x" } })
      expect(message(result)).toContain("isn't allowed")
      expect(message(result)).toContain(GRAPH_QUERY_HEADERS.join(", "))
      expect(fetchSpy).not.toHaveBeenCalled()
    },
  )

  it.each([
    ["a line break", "a\nb"],
    ["a DEL", "a\u007f"],
    ["a character above U+00FF", "aĀ"],
    ["a blank value", "  "],
  ])("refuses a value with %s", async (_, value) => {
    const result = await graphQuery({ method: "GET", path: "/me", headers: { Prefer: value } })
    expect(message(result)).toContain("header Prefer is blank or holds a character")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("refuses the same header twice in different case, which fetch would join", async () => {
    const result = await graphQuery({ method: "GET", path: "/me", headers: { "If-Match": "a", "if-match": "b" } })
    expect(message(result)).toContain("given twice")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("keeps a refused name on one line and never echoes a value", async () => {
    const name = await graphQuery({ method: "GET", path: "/me", headers: { "X\nForged: line": "v" } })
    expect(message(name).split("\n")).toHaveLength(1)
    const value = await graphQuery({ method: "GET", path: "/me", headers: { Prefer: "secret\nForged: line" } })
    expect(message(value)).not.toContain("secret")
  })

  it("passes every allowed header through, in any case", async () => {
    const headers = { "if-match": 'W/"1"', Prefer: "return=minimal", ConsistencyLevel: "eventual" }
    const result = await graphQuery({ method: "PATCH", path: "/planner/tasks/t1", body: "{}", headers })
    expect(result.isRight()).toBe(true)
    const sent = new Headers(fetchSpy.mock.calls[0]?.[1]?.headers)
    expect(sent.get("if-match")).toBe('W/"1"')
    expect(sent.get("prefer")).toBe("return=minimal")
    expect(sent.get("authorization")).toBe("Bearer TOKEN")
  })
})

describe("upload tools", () => {
  const path = "/me/drive/root:/a.txt:/content"

  it.each([
    ["an if_match with a line break", { if_match: '"{A},1"\nX: y' }, "if_match holds characters"],
    ["a content_type with a line break", { content_type: "text/plain\nX: y" }, "content_type must be a MIME type"],
  ])("upload_file refuses %s", async (_, extra, expected) => {
    const result = await uploadFile({ path, content: "x", ...extra })
    expect(message(result)).toContain(expected)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it.each([
    ["an if_match with a line break", { if_match: '"{A},1"\nX: y' }, "if_match holds characters"],
    ["a content_type with a line break", { content_type: "text/plain\nX: y" }, "content_type must be a MIME type"],
  ])("upload_file_from_path refuses %s before reading the file", async (_, extra, expected) => {
    const result = await uploadFileFromPath({ local_path: "/no/such/file", path, ...extra })
    expect(message(result)).toContain(expected)
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})

describe("update_planner_task etag", () => {
  it.each([
    ["a line break", 'W/"1"\nX: y'],
    ["only spaces", "   "],
  ])("refuses an etag with %s", async (_, etag) => {
    const result = await updatePlannerTask({ task_id: "t1", etag, title: "x" })
    expect(message(result)).toContain("etag is blank or holds characters")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("still looks the etag up when it is left out", async () => {
    fetchSpy
      .mockResolvedValueOnce(json({ "@odata.etag": 'W/"1"' }))
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
    const result = await updatePlannerTask({ task_id: "t1", title: "x" })
    expect(result.isRight()).toBe(true)
    expect(new Headers(fetchSpy.mock.calls[1]?.[1]?.headers).get("if-match")).toBe('W/"1"')
  })
})

describe("set_access_token", () => {
  it.each([
    ["a space", "a.b c"],
    ["a line break", "a.b\nc"],
    ["a non-ASCII character", "a.b.é"],
    ["nothing", "   "],
  ])("refuses a token with %s, which would break every later request", (_, token) => {
    const result = setAccessTokenTool({ access_token: token })
    expect(message(result)).toContain("access_token holds")
    expect(message(result)).not.toContain("a.b")
  })
})
