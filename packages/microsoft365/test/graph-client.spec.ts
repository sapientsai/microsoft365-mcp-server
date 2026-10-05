import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { getGraphClient, initializeGraphClient } from "../src/client/graph-client"

// Phase 2b: graph-client no longer imports the server auth module — it receives an
// AuthStrategy. These tests lock that seam.
describe("graph-client AuthStrategy injection", () => {
  afterEach(() => vi.unstubAllGlobals())

  const stubFetch = (json: unknown, ok = true, status = 200) =>
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve({
          ok,
          status,
          text: () => Promise.resolve(JSON.stringify(json)),
          json: () => Promise.resolve(json),
          headers: new Headers(),
        } as Response),
      ),
    )

  it("uses the injected strategy's token as the Bearer credential", async () => {
    const getAccessToken = vi.fn(() => Promise.resolve(Right("INJECTED.TOKEN")))
    const auth: AuthStrategy = { getAccessToken }
    stubFetch({ value: [] })

    const client = initializeGraphClient(auth)
    await client.listMessages()

    expect(getAccessToken).toHaveBeenCalledOnce()
    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer INJECTED.TOKEN")
  })

  // The mail-tools spec mocks the client, so it only proves "text" was passed along. The header
  // string itself is what Graph parses, and a malformed one is silently ignored — the body comes
  // back as HTML and nothing errors. Verified against live Graph v1.0: this exact string returns
  // body.contentType "text".
  it("sends the Prefer header verbatim when a body format is requested", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubFetch({ id: "m1" })

    const client = initializeGraphClient(auth)
    await client.getMessage("m1", "text")

    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Prefer).toBe('outlook.body-content-type="text"')
  })

  it("sends no Prefer header when no body format is requested", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubFetch({ id: "m1" })

    const client = initializeGraphClient(auth)
    await client.getMessage("m1")

    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>).Prefer).toBeUndefined()
  })

  it("short-circuits to an auth error when the strategy fails (fetch never called)", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Left({ type: "token", message: "no token" })) }
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)

    const client = initializeGraphClient(auth)
    const result = await client.getMe()

    expect(result.isLeft()).toBe(true)
    expect((result.value as { type: string }).type).toBe("auth")
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("initializeGraphClient registers the client as the active singleton", () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    initializeGraphClient(auth)
    expect(getGraphClient().isNone()).toBe(false)
  })

  it("graphQuery forwards caller-supplied headers (e.g. If-Match) onto the request", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    stubFetch({ ok: true })

    const client = initializeGraphClient(auth)
    await client.graphQuery("PATCH", "/planner/tasks/abc/details", { description: "x" }, undefined, {
      "If-Match": 'W/"etag123"',
    })

    const [, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect((init.headers as Record<string, string>)["If-Match"]).toBe('W/"etag123"')
  })

  it("updatePlannerTaskDetails sends the details path with the If-Match ETag", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }
    stubFetch({ ok: true })

    const client = initializeGraphClient(auth)
    await client.updatePlannerTaskDetails("abc", { description: "hi" }, 'W/"e"')

    const [url, init] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    expect(url).toContain("/planner/tasks/abc/details")
    expect(init.method).toBe("PATCH")
    expect((init.headers as Record<string, string>)["If-Match"]).toBe('W/"e"')
  })

  describe("connector report fixes: request shapes", () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    const firstCall = () => vi.mocked(fetch).mock.calls[0] as [string, RequestInit]

    it("sends the same text Prefer header for an event as for a message", async () => {
      stubFetch({ id: "e1" })
      await initializeGraphClient(auth).getEvent("e1", "text")
      expect((firstCall()[1].headers as Record<string, string>).Prefer).toBe('outlook.body-content-type="text"')
    })

    it("sends no Prefer header for an event when no format is given", async () => {
      stubFetch({ id: "e1" })
      await initializeGraphClient(auth).getEvent("e1")
      expect((firstCall()[1].headers as Record<string, string>).Prefer).toBeUndefined()
    })

    it("limits a OneDrive search with $top", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).searchFiles("ONC", { $top: 25 })
      expect(decodeURIComponent(firstCall()[0])).toContain("/me/drive/root/search(q='ONC')?$top=25")
    })

    // q is an OData string literal; an undoubled apostrophe ends it early and Graph rejects the request.
    it("doubles an apostrophe in a search term", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).searchFiles("O'Brien notes")
      expect(decodeURIComponent(firstCall()[0])).toContain("/me/drive/root/search(q='O''Brien notes')")

      vi.unstubAllGlobals()
      stubFetch({ value: [] })
      await initializeGraphClient(auth).searchSiteFiles("s1", "it's")
      expect(decodeURIComponent(firstCall()[0])).toContain("/sites/s1/drive/root/search(q='it''s')")
    })

    it("limits a SharePoint search with $top, in the default library or a named drive", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).searchSiteFiles("s1", "Annual", undefined, { $top: 10 })
      expect(decodeURIComponent(firstCall()[0])).toContain("/sites/s1/drive/root/search(q='Annual')?$top=10")

      vi.unstubAllGlobals()
      stubFetch({ value: [] })
      await initializeGraphClient(auth).searchSiteFiles("s1", "Annual", "d1", { $top: 10 })
      expect(decodeURIComponent(firstCall()[0])).toContain("/sites/s1/drives/d1/root/search(q='Annual')?$top=10")
    })

    it("asks for chat members and the last message, newest first", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).listChats({
        $expand: ["members", "lastMessagePreview"],
        $orderby: "lastMessagePreview/createdDateTime desc",
        $top: 5,
      })
      const url = decodeURIComponent(firstCall()[0])
      expect(url).toContain("$expand=members,lastMessagePreview")
      expect(url).toContain("$orderby=lastMessagePreview/createdDateTime desc")
      expect(url).toContain("$top=5")
    })
  })

  describe("message listing paths", () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    const requestedUrl = () => (vi.mocked(fetch).mock.calls[0] as [string, RequestInit])[0]

    it("lists every folder when no folder is given", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).listMessages()
      expect(requestedUrl()).toMatch(/\/v1\.0\/me\/messages(\?|$)/)
    })

    it("lists one folder by well-known name", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).listMessages(undefined, "inbox")
      expect(requestedUrl()).toContain("/me/mailFolders/inbox/messages")
    })

    it("encodes a folder ID so its characters cannot change the path", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).listMessages(undefined, "AAMk/AB+c=")
      expect(requestedUrl()).toContain("/me/mailFolders/AAMk%2FAB%2Bc%3D/messages")
    })

    it("pages through the same folder path", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).listAllMessages(undefined, "sentitems")
      expect(requestedUrl()).toContain("/me/mailFolders/sentitems/messages")
    })
  })
})
