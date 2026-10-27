import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { getGraphClient, initializeGraphClient } from "../src/client/graph-client"

// Phase 2b: graph-client no longer imports the server auth module — it receives an
// AuthStrategy. These tests lock that seam.
describe("graph-client AuthStrategy injection", () => {
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.unstubAllEnvs()
  })

  it.each(["download", "upload"])("honours retry settings for raw %s content", async (operation) => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "1")
    vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
    vi.stubEnv("MS365_GRAPH_VERSION", "beta")
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(new Response("throttled", { status: 429 }))
      .mockResolvedValueOnce(new Response(JSON.stringify({ id: "file" })))
    vi.stubGlobal("fetch", fetchSpy)
    const client = initializeGraphClient({ getAccessToken: async () => Right("TOKEN") })
    const result =
      operation === "download"
        ? await client.downloadFileContent("file")
        : await client.uploadFile("/me/drive/root:/x.txt:/content", "text")
    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(2)
    expect(fetchSpy.mock.calls[0]?.[0]).toContain("/beta/me/drive/")
  })

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

  // Routes each request by URL: the attachment list, then one read per cloud link.
  const stubAttachmentFetch = (list: unknown, links: Record<string, unknown | undefined>) =>
    vi.stubGlobal(
      "fetch",
      vi.fn((url: string) => {
        const path = new URL(url).pathname
        const linkId = /\/attachments\/([^/]+)$/.exec(path)?.[1]
        const json = linkId === undefined ? list : links[decodeURIComponent(linkId)]
        const ok = json !== undefined
        return Promise.resolve({
          ok,
          status: ok ? 200 : 404,
          text: () => Promise.resolve(JSON.stringify(json ?? { error: { message: "not found" } })),
          json: () => Promise.resolve(json),
          headers: new Headers({ "content-type": "application/json" }),
        } as Response)
      }),
    )

  const FILE = { "@odata.type": "#microsoft.graph.fileAttachment", id: "FILE", name: "invoice.pdf", size: 1024 }
  const LINK = { "@odata.type": "#microsoft.graph.referenceAttachment", id: "LINK/1=", name: "Renovation invoices" }

  // A fileAttachment carries its whole base64 payload unless a $select leaves contentBytes out, and
  // Graph rejects a $select naming referenceAttachment-only properties for every message:
  //   Parsing OData Select and Expand failed: Could not find a property named 'sourceUrl'
  //   on type 'microsoft.graph.attachment'.
  // So the list selects the base properties, and only those.
  it("lists attachments with a base-type $select that leaves out contentBytes and derived properties", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubAttachmentFetch({ value: [FILE] }, {})

    const client = initializeGraphClient(auth)
    await client.listAttachments("MSG-ID")

    expect(vi.mocked(fetch)).toHaveBeenCalledOnce()
    const [url] = vi.mocked(fetch).mock.calls[0] as unknown as [string]
    const select = new URL(url).searchParams.get("$select")?.split(",") ?? []
    expect(select).toEqual(expect.arrayContaining(["id", "name", "contentType", "size"]))
    for (const excluded of ["contentBytes", "sourceUrl", "providerType", "permission", "isFolder"]) {
      expect(select).not.toContain(excluded)
    }
  })

  it("reads each cloud link on its own to get its URL, and no file attachment", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubAttachmentFetch(
      { value: [FILE, LINK] },
      {
        "LINK/1=": {
          ...LINK,
          sourceUrl: "https://www.icloud.com/iclouddrive/EXAMPLE",
          providerType: "other",
          permission: "view",
          isFolder: true,
        },
      },
    )

    const client = initializeGraphClient(auth)
    const result = await client.listAttachments("MSG-ID")

    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(2)
    const [linkUrl] = vi.mocked(fetch).mock.calls[1] as unknown as [string]
    expect(new URL(linkUrl).pathname).toMatch(/\/me\/messages\/MSG-ID\/attachments\/LINK%2F1%3D$/)

    const attachments = (result.value as { value: ReadonlyArray<Record<string, unknown>> }).value
    expect(attachments.find((a) => a.id === "FILE")).toEqual(FILE)
    expect(attachments.find((a) => a.id === "LINK/1=")).toMatchObject({
      "@odata.type": "#microsoft.graph.referenceAttachment",
      sourceUrl: "https://www.icloud.com/iclouddrive/EXAMPLE",
      providerType: "other",
      permission: "view",
      isFolder: true,
    })
  })

  it("still lists a cloud link, without its URL, when reading it fails", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubAttachmentFetch({ value: [LINK] }, {})

    const client = initializeGraphClient(auth)
    const result = await client.listAttachments("MSG-ID")

    expect(result.isRight()).toBe(true)
    const [link] = (result.value as { value: ReadonlyArray<Record<string, unknown>> }).value
    expect(link).toMatchObject({ id: "LINK/1=", name: "Renovation invoices" })
    expect(link?.sourceUrl).toBeUndefined()
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

  describe("chat member lookup", () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    const requestedUrl = () => (vi.mocked(fetch).mock.calls[0] as [string, RequestInit])[0]

    // A guest's sign-in name contains "#EXT#"; unencoded, "#" ends the path and the lookup 404s.
    it("encodes # in a guest's sign-in name", async () => {
      stubFetch({ id: "guest-id" })
      await initializeGraphClient(auth).getUserByKey("pat_vendor.com#EXT#@civala.onmicrosoft.com")
      expect(requestedUrl()).toContain("/users/pat_vendor.com%23EXT%23@civala.onmicrosoft.com?")
      expect(decodeURIComponent(requestedUrl())).toContain("$select=id,displayName,mail,userPrincipalName,userType")
    })

    // userType comes back only when selected, and the caller's guest status decides their chat role.
    it("selects userType when reading the caller for a chat", async () => {
      stubFetch({ id: "me-id" })
      await initializeGraphClient(auth).getMeAsChatMember()
      expect(requestedUrl()).toContain("/me?")
      expect(decodeURIComponent(requestedUrl())).toContain("$select=id,displayName,mail,userPrincipalName,userType")
    })

    it("doubles a quote in the mail filter", async () => {
      stubFetch({ value: [] })
      await initializeGraphClient(auth).findUsersByMail("o'brien@civala.com")
      expect(decodeURIComponent(requestedUrl())).toContain("$filter=mail eq 'o''brien@civala.com'")
    })
  })
})

// Graph does not echo the request's mailbox segment back verbatim. A live delegated scan of
// /users/{address with %40}/... returned next links with a literal "@", and a folder scan's link
// in OData key form, mailFolders('inbox'). These are those shapes with the address swapped.
describe("nextMessageScanPage in another mailbox", () => {
  const PREFIX = "/users/jake%40example.com"

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  const followWith = async (nextLink: string) => {
    const fetchSpy = vi.fn(() =>
      Promise.resolve(new Response(JSON.stringify({ value: [] }), { headers: { "Content-Type": "application/json" } })),
    )
    vi.stubGlobal("fetch", fetchSpy)
    const client = initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("TOKEN")) })
    return { result: await client.nextMessageScanPage(nextLink, PREFIX), fetchSpy }
  }

  it.each([
    "https://graph.microsoft.com/v1.0/users/jake@example.com/mailFolders('inbox')/messages?%24select=id%2creceivedDateTime&%24filter=receivedDateTime+ge+2026-01-01T00%3a00%3a00Z&%24orderby=receivedDateTime+desc&%24top=2&%24skip=4",
    "https://graph.microsoft.com/v1.0/users/jake@example.com/messages?%24select=id&%24filter=receivedDateTime+ge+2026-10-01T00%3a00%3a00Z&%24orderby=receivedDateTime+desc&%24top=2&%24skip=2",
  ])("follows the link Graph returned for a %%40-encoded scan: %s", async (nextLink) => {
    const { result, fetchSpy } = await followWith(nextLink)

    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledOnce()
    expect(fetchSpy.mock.calls[0]![0]).toBe(nextLink)
  })

  it.each([
    "https://graph.microsoft.com/v1.0/users/other@example.com/mailFolders('inbox')/messages?%24skip=4",
    "https://graph.microsoft.com/v1.0/users/other%40example.com/messages?%24skip=2",
    "https://graph.microsoft.com/v1.0/me/messages?%24skip=2",
  ])("refuses a link naming a different mailbox: %s", async (nextLink) => {
    const { result, fetchSpy } = await followWith(nextLink)

    expect(result.isLeft()).toBe(true)
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  // decodeURIComponent throws on a stray "%"; the link must be refused, not crash the scan.
  it.each([
    "https://graph.microsoft.com/v1.0/users/jake%40example.com%/messages?%24skip=2",
    "https://graph.microsoft.com/v1.0/users('jake%zz')/messages?%24skip=2",
  ])("refuses a link with a malformed escape instead of throwing: %s", async (nextLink) => {
    const { result, fetchSpy } = await followWith(nextLink)

    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toBe("Invalid Graph message continuation URL.")
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})
