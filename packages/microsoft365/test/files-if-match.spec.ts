import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Right } from "functype/either"
import { afterAll, afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth/token-context", () => ({ getContextToken: vi.fn(() => "GRAPH.TOKEN") }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))

import { initializeGraphClient } from "../src/client/graph-client"
import { getUploadConfig, uploadFile, uploadFileFromPath } from "../src/tools/files-tools"
import { formatDriveItemDetail, formatDriveItemSummary } from "../src/utils/formatters"

// Issue 187: a read-modify-write must be able to say "only if nobody changed it since I read it".
const ETAG = '"{5A1F},3"'
const json = (value: unknown, status = 200) =>
  new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } })
const item = { id: "file-1", name: "state.json", eTag: '"{5A1F},4"', file: { mimeType: "application/json" } }
const stale = () => json({ error: { code: "notAllowed", message: "" } }, 412)
const sessionThenItem = () =>
  vi
    .fn()
    .mockResolvedValueOnce(json({ uploadUrl: "https://upload.example/session" }))
    .mockResolvedValueOnce(json(item))

beforeEach(() => {
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
  initializeGraphClient({ getAccessToken: async () => Right("TOKEN") })
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("eTag on reads", () => {
  it("shows the eTag in a drive item's detail and in a list line", () => {
    expect(formatDriveItemDetail(item)).toContain(`- eTag: ${ETAG.replace(",3", ",4")}`)
    expect(formatDriveItemSummary(item)).toContain(`eTag ${item.eTag}`)
  })

  it("leaves the eTag out when Graph didn't send one", () => {
    expect(formatDriveItemDetail({ id: "x", name: "x" })).not.toContain("eTag")
  })
})

describe("upload_file with if_match", () => {
  it("writes through an upload session with If-Match, and returns the new eTag", async () => {
    const fetchSpy = sessionThenItem()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFile({
      path: "/me/drive/root:/state.json:/content",
      content: '{"n":2}',
      content_type: "application/json",
      if_match: ETAG,
    })
    expect(result.isRight()).toBe(true)
    expect(result.value).toContain(`- eTag: ${item.eTag}`)
    const [url, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect(url).toContain(":/createUploadSession")
    expect(init.headers).toMatchObject({ "If-Match": ETAG })
  })

  it("reports a stale eTag as precondition-failed and writes nothing", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(stale()))
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFile({ path: "/me/drive/root:/state.json:/content", content: "x", if_match: ETAG })
    expect(result.isLeft()).toBe(true)
    expect((result.value as Error).message).toContain("precondition-failed")
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("refuses if_match with conflict_behavior rename, without calling Graph", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFile({
      path: "/me/drive/root:/state.json:/content",
      content: "x",
      conflict_behavior: "rename",
      if_match: ETAG,
    })
    expect((result.value as Error).message).toContain('not "rename"')
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("refuses a blank if_match instead of writing unconditionally", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFile({ path: "/me/drive/root:/state.json:/content", content: "x", if_match: "  " })
    expect((result.value as Error).message).toContain("if_match is blank")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("writes conditionally to an item-ID path through /items/{id}/createUploadSession", async () => {
    const fetchSpy = sessionThenItem()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFile({ path: "/me/drive/items/file-1/content", content: "x", if_match: ETAG })
    expect(result.isRight()).toBe(true)
    const [url, init] = fetchSpy.mock.calls[0] as [string, RequestInit]
    expect(url).toMatch(/\/me\/drive\/items\/file-1\/createUploadSession$/)
    expect(init.headers).toMatchObject({ "If-Match": ETAG })
  })

  it("without if_match, keeps the single PUT", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(json(item)))
    vi.stubGlobal("fetch", fetchSpy)
    expect((await uploadFile({ path: "/me/drive/root:/state.json:/content", content: "x" })).isRight()).toBe(true)
    const [url, init] = fetchSpy.mock.calls[0] as unknown as [string, RequestInit]
    expect(init.method).toBe("PUT")
    expect(url).toContain("conflictBehavior=rename")
  })
})

describe("upload_file_from_path with if_match", () => {
  const dir = mkdtempSync(join(tmpdir(), "if-match-spec-"))
  const local = join(dir, "state.json")
  writeFileSync(local, '{"n":2}')
  afterAll(() => rmSync(dir, { recursive: true, force: true }))

  it("sends If-Match on the upload session", async () => {
    const fetchSpy = sessionThenItem()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadFileFromPath({
      local_path: local,
      path: "/me/drive/root:/state.json:/content",
      if_match: ETAG,
    })
    expect(result.isRight()).toBe(true)
    expect((fetchSpy.mock.calls[0] as [string, RequestInit])[1].headers).toMatchObject({ "If-Match": ETAG })
  })

  it("reports a stale eTag as precondition-failed", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve(stale())),
    )
    const result = await uploadFileFromPath({
      local_path: local,
      path: "/me/drive/root:/state.json:/content",
      if_match: ETAG,
    })
    expect((result.value as Error).message).toContain("precondition-failed")
  })
})

describe("get_upload_config with ifMatch", () => {
  beforeEach(() => {
    vi.stubEnv("MS365_PUBLIC_BASE_URL", "https://ms365.example.com")
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve({ status: 401 } as Response)),
    )
  })

  it("puts ifMatch on the relay URL and implies replace", async () => {
    const result = await getUploadConfig({ path: "/me/drive/root:/state.json:/content", ifMatch: ETAG })
    expect(result.isRight()).toBe(true)
    const payload = JSON.parse(result.value as string) as {
      uploadUrl: string
      conflictBehavior: string
      ifMatch: string
    }
    const url = new URL(payload.uploadUrl)
    expect(url.searchParams.get("ifMatch")).toBe(ETAG)
    expect(url.searchParams.get("conflictBehavior")).toBe("replace")
    expect(payload.ifMatch).toBe(ETAG)
  })

  it("refuses a blank ifMatch, which the relay would read as no condition", async () => {
    const result = await getUploadConfig({ path: "/me/drive/root:/state.json:/content", ifMatch: " " })
    expect((result.value as Error).message).toContain("ifMatch is blank")
  })

  it("refuses ifMatch with conflictBehavior fail", async () => {
    const result = await getUploadConfig({
      path: "/me/drive/root:/state.json:/content",
      ifMatch: ETAG,
      conflictBehavior: "fail",
    })
    expect((result.value as Error).message).toContain('not "fail"')
  })
})
