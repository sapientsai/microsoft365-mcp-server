import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import { getMessage, scanMessages } from "../src/tools/mail-tools"

const auth = { getAccessToken: () => Promise.resolve(Right("test-token")) }
const message = (id: string) => ({ id, subject: id, hasAttachments: true })
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } })
const cursorFrom = (output: unknown) => {
  const match = /cursor: "([^"]+)"/.exec(String(output))
  expect(match).not.toBeNull()
  return match![1]!
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe("scan pagination through the Graph client", () => {
  it.each([{ filter: "hasAttachments eq true" }, { search: "invoice" }])(
    "follows every nextLink without dropping messages for %j",
    async (query) => {
      const nextLink =
        "https://graph.microsoft.com/v1.0/me/mailFolders/archive/messages?$skiptoken=opaque%2Btoken%3D&$top=2"
      const thirdLink = "https://graph.microsoft.com/v1.0/me/mailFolders/archive/messages?$skiptoken=last%2Fpage"
      const fetchSpy = vi
        .fn()
        .mockResolvedValueOnce(json({ value: [message("a")], "@odata.nextLink": nextLink }))
        .mockResolvedValueOnce(json({ value: [], "@odata.nextLink": thirdLink }))
        .mockResolvedValueOnce(json({ value: [message("b"), message("c")] }))
      vi.stubGlobal("fetch", fetchSpy)
      initializeGraphClient(auth)

      const first = await scanMessages({ ...query, folder: "archive", top: 2 })
      expect(first.isRight()).toBe(true)
      expect(first.value).toContain("INCOMPLETE")
      const firstUrl = new URL(fetchSpy.mock.calls[0]![0] as string)
      expect(firstUrl.pathname).toBe("/v1.0/me/mailFolders/archive/messages")
      expect(firstUrl.searchParams.get("$top")).toBe("2")
      expect(firstUrl.searchParams.has("$skip")).toBe(false)
      expect(firstUrl.searchParams.has("$orderby")).toBe(false)

      const second = await scanMessages({ cursor: cursorFrom(first.value) })
      expect(second.value).toContain("INCOMPLETE")
      expect(fetchSpy.mock.calls[1]![0]).toBe(nextLink)
      const third = await scanMessages({ cursor: cursorFrom(second.value) })
      expect(fetchSpy.mock.calls[2]![0]).toBe(thirdLink)
      expect(third.value).toContain("|b|")
      expect(third.value).toContain("|c|")
      expect(third.value).not.toContain("INCOMPLETE")
      expect(third.value).toContain("2 in archive")
      if ("search" in query) expect(third.value).toContain("1,000")
      for (const call of fetchSpy.mock.calls) {
        expect((call[1] as RequestInit).headers).toMatchObject({ Authorization: "Bearer test-token" })
      }
    },
  )

  // The shared request layer retries some failures on reads by itself, so a retried status
  // would succeed on the retry and never reach scanMessages as a failure. A 500 isn't retried
  // today; turning retries off as well keeps the test valid if that list ever changes.
  it("keeps a failed continuation retryable", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$skip=37"
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ value: [], "@odata.nextLink": nextLink }))
      .mockResolvedValueOnce(json({ error: { message: "try again" } }, 500))
      .mockResolvedValueOnce(json({ value: [] }))
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    const first = await scanMessages({})
    const cursor = cursorFrom(first.value)
    expect((await scanMessages({ cursor })).isLeft()).toBe(true)
    expect((await scanMessages({ cursor })).isRight()).toBe(true)
    expect(fetchSpy.mock.calls[1]![0]).toBe(nextLink)
    expect(fetchSpy.mock.calls[2]![0]).toBe(nextLink)
    expect(fetchSpy).toHaveBeenCalledTimes(3)
  })

  it("rejects unknown cursors and parameter changes without making a request", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    expect((await scanMessages({ cursor: "unknown" })).isLeft()).toBe(true)
    expect((await scanMessages({ cursor: "unknown", top: 2 })).isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("does not reuse a cursor after the active client changes", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValue(
        json({ value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=10" }),
      )
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    const first = await scanMessages({})
    initializeGraphClient(auth)
    expect((await scanMessages({ cursor: cursorFrom(first.value) })).isLeft()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledOnce()
  })

  it.each([0, -1, 1001, 1.5, NaN])("rejects invalid page size %s", async (top) => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    expect((await scanMessages({ top })).isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("accepts the maximum page size and resolves a returned message ref", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(json({ value: [message("graph-id")] }))
      .mockResolvedValueOnce(json(message("graph-id")))
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    const result = await scanMessages({ top: 1000 })
    expect(new URL(fetchSpy.mock.calls[0]![0] as string).searchParams.get("$top")).toBe("1000")
    const ref = /\n(\d+)\|/.exec(String(result.value))![1]!
    expect((await getMessage({ message_id: ref })).isRight()).toBe(true)
    expect(fetchSpy.mock.calls[1]![0]).toBe("https://graph.microsoft.com/v1.0/me/messages/graph-id")
  })

  it.each([
    "not a URL",
    "https://evil.example/v1.0/me/messages?$skip=2",
    "https://graph.microsoft.com/v1.0/me/sendMail",
    "https://graph.microsoft.com/v1.0/users/someone/messages",
  ])("rejects an unexpected continuation URL %s", async (nextLink) => {
    const fetchSpy = vi.fn().mockResolvedValue(json({ value: [], "@odata.nextLink": nextLink }))
    vi.stubGlobal("fetch", fetchSpy)
    initializeGraphClient(auth)
    const first = await scanMessages({})
    expect((await scanMessages({ cursor: cursorFrom(first.value) })).isLeft()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledOnce()
  })
})
