import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import {
  ifMatchConflictProblem,
  PRECONDITION_FAILED_MESSAGE,
  sessionUpload,
  uploadDriveContent,
} from "../src/upload/upload"

const apiBase = "https://graph.microsoft.com/v1.0"
const path = "/me/drive/root:/state.json:/content"
const data = Buffer.from('{"n":2}')
const item = { id: "file-1", name: "state.json", eTag: '"{ABC},4"' }
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status })
const ETAG = '"{ABC},3"'

describe("conditional drive writes (If-Match)", () => {
  beforeEach(() => {
    vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
    vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  })
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.unstubAllEnvs()
  })

  it("without ifMatch, keeps today's single PUT for small content", async () => {
    const fetchSpy = vi.fn().mockResolvedValueOnce(json(item))
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "token", data, "application/json", "rename")
    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    const [url, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect(init.method).toBe("PUT")
    expect(url).toContain("@microsoft.graph.conflictBehavior=rename")
    expect(init.headers).not.toHaveProperty("If-Match")
  })

  it("with ifMatch, sends If-Match on createUploadSession and replaces, even for small content", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ uploadUrl: "https://upload.example/session" }))
      .mockResolvedValueOnce(json(item, 200))
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "token", data, "application/json", "replace", {
      ifMatch: ETAG,
    })
    expect(result.isRight()).toBe(true)
    expect(result.value).toMatchObject({ eTag: '"{ABC},4"' })
    const [sessionUrl, sessionInit] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect(sessionUrl).toBe(`${apiBase}/me/drive/root:/state.json:/createUploadSession`)
    expect(sessionInit.method).toBe("POST")
    expect(sessionInit.headers).toMatchObject({ "If-Match": ETAG })
    expect(JSON.parse(sessionInit.body as string)).toEqual({ item: { "@microsoft.graph.conflictBehavior": "replace" } })
    expect((fetchSpy.mock.calls[1] as [string, RequestInit])[1].method).toBe("PUT")
  })

  it("reports a stale eTag as precondition-failed, sends no bytes and doesn't retry", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "3")
    const fetchSpy = vi
      .fn()
      .mockResolvedValue(json({ error: { code: "notAllowed", message: "ETag does not match" } }, 412))
    vi.stubGlobal("fetch", fetchSpy)
    const result = await sessionUpload(apiBase, path, "token", data, "replace", ETAG)
    expect(result.isLeft()).toBe(true)
    expect(result.value).toEqual({ type: "api", status: 412, message: PRECONDITION_FAILED_MESSAGE })
    expect(PRECONDITION_FAILED_MESSAGE).toMatch(/^precondition-failed: /)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("answers a 412 on the committing chunk as precondition-failed too", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ uploadUrl: "https://upload.example/session" }))
      .mockResolvedValueOnce(json({ error: { code: "notAllowed", message: "" } }, 412))
      .mockResolvedValueOnce(new Response(null, { status: 204 }))
    vi.stubGlobal("fetch", fetchSpy)
    const result = await sessionUpload(apiBase, path, "token", data, "replace", ETAG)
    expect(result.value).toEqual({ type: "api", status: 412, message: PRECONDITION_FAILED_MESSAGE })
  })

  it("refuses a path with no /content form instead of posting to it", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await sessionUpload(apiBase, "/me/drive/items/file-1", "token", data, "replace", ETAG)
    expect((result.value as { message: string }).message).toContain('must end with ":/content"')
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("refuses a blank ifMatch without calling Graph", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "token", data, "text/plain", "replace", { ifMatch: " " })
    expect((result.value as { message: string }).message).toContain("if_match is blank")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("refuses empty content with ifMatch without calling Graph", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "token", Buffer.alloc(0), "text/plain", "replace", {
      ifMatch: ETAG,
    })
    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("allows ifMatch only with replace or no explicit conflict behavior", () => {
    expect(ifMatchConflictProblem(undefined, ETAG)).toBeUndefined()
    expect(ifMatchConflictProblem("replace", ETAG)).toBeUndefined()
    expect(ifMatchConflictProblem("rename", undefined)).toBeUndefined()
    expect(ifMatchConflictProblem("rename", ETAG)).toContain('not "rename"')
    expect(ifMatchConflictProblem("fail", ETAG)).toContain('not "fail"')
  })
})
