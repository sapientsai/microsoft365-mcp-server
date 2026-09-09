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

  // Graph validates $select for /attachments against the BASE attachment type. sourceUrl and
  // friends live on the derived referenceAttachment, so naming them in a $select makes Graph
  // reject the whole request — for every message, not only ones carrying a cloud link:
  //
  //   Parsing OData Select and Expand failed: Could not find a property named 'sourceUrl'
  //   on type 'microsoft.graph.attachment'.
  //
  // That shipped once and broke list_attachments and save_attachment outright. The previous
  // version of this test asserted the broken behaviour, because it only checked the URL we build
  // and never what Graph does with it. Assert the absence instead — it is the property that
  // actually keeps the endpoint working.
  it("does not $select derived referenceAttachment properties (Graph rejects the request)", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubFetch({ value: [] })

    const client = initializeGraphClient(auth)
    await client.listAttachments("MSG-ID")

    const [url] = vi.mocked(fetch).mock.calls[0] as [string, RequestInit]
    const requested = decodeURIComponent(url)
    for (const derived of ["sourceUrl", "providerType", "permission", "isFolder"]) {
      expect(requested).not.toContain(derived)
    }
  })

  // The reference fields and @odata.type must still reach the caller — dropping them is how cloud
  // links became invisible in the first place. With no $select, Graph returns them itself.
  it("passes through @odata.type and reference fields, and strips contentBytes", async () => {
    const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
    stubFetch({
      value: [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          id: "FILE",
          name: "invoice.pdf",
          size: 1024,
          contentBytes: "QUJD".repeat(10_000),
        },
        {
          "@odata.type": "#microsoft.graph.referenceAttachment",
          id: "LINK",
          name: "Renovation invoices",
          sourceUrl: "https://www.icloud.com/iclouddrive/EXAMPLE",
          providerType: "other",
          permission: "view",
          isFolder: true,
        },
      ],
    })

    const client = initializeGraphClient(auth)
    const result = await client.listAttachments("MSG-ID")

    expect(result.isRight()).toBe(true)
    const attachments = (result.value as { value: ReadonlyArray<Record<string, unknown>> }).value

    const file = attachments.find((a) => a.id === "FILE")
    expect(file?.["@odata.type"]).toBe("#microsoft.graph.fileAttachment")
    // A fileAttachment's base64 payload would be megabytes of binary through the model.
    expect(file).not.toHaveProperty("contentBytes")
    expect(file?.name).toBe("invoice.pdf")

    const link = attachments.find((a) => a.id === "LINK")
    expect(link?.sourceUrl).toBe("https://www.icloud.com/iclouddrive/EXAMPLE")
    expect(link?.providerType).toBe("other")
    expect(link?.isFolder).toBe(true)
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
})
