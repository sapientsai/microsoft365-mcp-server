import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const server = vi.hoisted(() => ({
  addTool: vi.fn(),
  start: vi.fn(),
  post: vi.fn(),
  put: vi.fn(),
  instructions: "",
}))
vi.mock("fastmcp", async (original) => ({
  ...(await original<typeof import("fastmcp")>()),
  FastMCP: class {
    constructor(options: { instructions: string }) {
      server.instructions = options.instructions
    }
    addTool = server.addTool
    start = server.start
    getApp = () => ({ post: server.post, put: server.put })
  },
}))
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  initializeAuth: vi.fn(async () => Right(undefined)),
  getAccessToken: vi.fn(async () => Right("TOKEN")),
}))
vi.mock("../src/auth/oauth-provider", () => ({ createAzureAuthProvider: vi.fn() }))

beforeEach(() => {
  vi.resetModules()
  vi.clearAllMocks()
  for (const key of [
    "MS365_READ_ONLY",
    "MS365_PRESETS",
    "MS365_ENABLED_TOOLS",
    "MS365_REQUIRE_DRAFT",
    "MS365_ORG_MODE",
  ])
    vi.stubEnv(key, undefined)
  vi.stubEnv("TRANSPORT_TYPE", "httpStream")
})
afterEach(() => {
  vi.unstubAllEnvs()
  vi.unstubAllGlobals()
})

describe.each(["client-token", "oauth-proxy"])("%s server environment controls", (mode) => {
  it.each([
    ["MS365_READ_ONLY", "true"],
    ["MS365_PRESETS", "productivity"],
    ["MS365_ENABLED_TOOLS", "^list_"],
  ])("does not expose /upload when %s=%s", async (key, value) => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv(key, value)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(server.post).not.toHaveBeenCalledWith("/upload", expect.any(Function))
    expect(server.put).not.toHaveBeenCalledWith("/upload", expect.any(Function))
    expect(server.instructions).not.toContain("Upload workflows:")
    expect(server.addTool.mock.calls.map(([tool]) => tool.name)).not.toContain("get_upload_config")
  })

  it.each([undefined, "v1.0"])("uses the configured Graph version for /upload, with override %s", async (override) => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv("MS365_GRAPH_VERSION", "beta")
    vi.stubEnv("MS365_UPLOAD_TOKEN", "TOKEN")
    const fetchSpy = vi.fn().mockResolvedValue(new Response(JSON.stringify({ id: "uploaded" })))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    const handler = server.post.mock.calls[0]?.[1]
    const query: Record<string, string | undefined> = {
      path: "/me/drive/root:/test.txt:/content",
      apiVersion: override,
    }
    await handler({
      req: {
        query: (key: string) => query[key],
        header: () => "Bearer TOKEN",
        arrayBuffer: async () => new TextEncoder().encode("text").buffer,
      },
      json: vi.fn(),
    })
    expect(fetchSpy.mock.calls[0]?.[0]).toContain(`/${override ?? "beta"}/me/drive/`)
  })

  it("keeps /upload when get_upload_config is enabled", async () => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv("MS365_ENABLED_TOOLS", "^get_upload_config$")
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(server.post).toHaveBeenCalledWith("/upload", expect.any(Function))
    expect(server.put).toHaveBeenCalledWith("/upload", expect.any(Function))
  })
})

describe.each(["client-token", "oauth-proxy"])("%s draft attachment relay", (mode) => {
  const attachmentHandler = () => server.post.mock.calls.find(([path]) => path === "/upload/attachment")?.[1]
  const request = (query: Record<string, string | undefined>) => ({
    req: {
      query: (key: string) => query[key],
      header: () => "Bearer TOKEN",
      arrayBuffer: async () => new TextEncoder().encode(Buffer.from("%PDF").toString("base64")).buffer,
    },
    json: vi.fn(),
  })
  const graph = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status })

  beforeEach(() => {
    vi.stubEnv("MS365_AUTH_MODE", mode)
    vi.stubEnv("MS365_UPLOAD_TOKEN", "TOKEN")
    vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
  })

  it.each([
    ["MS365_READ_ONLY", "true"],
    ["MS365_PRESETS", "collaboration"],
    ["MS365_ENABLED_TOOLS", "^list_"],
  ])("does not expose /upload/attachment when %s=%s", async (key, value) => {
    vi.stubEnv(key, value)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(attachmentHandler()).toBeUndefined()
    expect(server.put).not.toHaveBeenCalledWith("/upload/attachment", expect.any(Function))
    expect(server.instructions).not.toContain("Mail attachments:")
    expect(server.addTool.mock.calls.map(([tool]) => tool.name)).not.toContain("get_attachment_upload_config")
  })

  it("mounts only /upload/attachment when only its tool is enabled, and keeps it under require-draft", async () => {
    vi.stubEnv("MS365_ENABLED_TOOLS", "^get_attachment_upload_config$")
    vi.stubEnv("MS365_REQUIRE_DRAFT", "true")
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())
    expect(server.post.mock.calls.map(([path]) => path)).toEqual(["/upload/attachment"])
    expect(server.put).toHaveBeenCalledWith("/upload/attachment", expect.any(Function))
    expect(server.instructions).toContain("Mail attachments:")
  })

  it("attaches the decoded file to a draft using the configured Graph version", async () => {
    vi.stubEnv("MS365_GRAPH_VERSION", "beta")
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(graph({ id: "draft", isDraft: true }))
      .mockResolvedValueOnce(graph({ id: "att", size: 4 }, 201))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    const context = request({ messageId: "draft", name: "report.pdf", encoding: "base64" })
    await attachmentHandler()?.(context)

    expect(context.json).toHaveBeenCalledWith({ id: "att", name: "report.pdf", size: 4 }, 200)
    expect(fetchSpy.mock.calls[1]?.[0]).toBe("https://graph.microsoft.com/beta/me/messages/draft/attachments")
    const body = JSON.parse(fetchSpy.mock.calls[1]?.[1]?.body as string) as {
      contentBytes: string
      contentType: string
    }
    expect(Buffer.from(body.contentBytes, "base64").toString()).toBe("%PDF")
    expect(body.contentType).toBe("application/pdf")
  })

  it("refuses a message that is not a draft, even with a valid upload credential", async () => {
    const fetchSpy = vi.fn().mockResolvedValueOnce(graph({ id: "sent", isDraft: false }))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    const context = request({ messageId: "sent", name: "report.pdf", encoding: "base64" })
    await attachmentHandler()?.(context)

    expect(context.json).toHaveBeenCalledWith({ error: expect.stringContaining("not a draft") }, 409)
    expect(fetchSpy).toHaveBeenCalledTimes(1)
  })

  it("refuses a bad credential with 401 before reading the body or attaching", async () => {
    // Shared-secret mode rejects the bearer itself; oauth-proxy mode learns it from Graph's 401.
    const fetchSpy = vi.fn().mockResolvedValue(graph({ error: { code: "InvalidAuthenticationToken" } }, 401))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    const context = request({ messageId: "draft", name: "report.pdf", encoding: "base64" })
    context.req.header = () => "Bearer wrong"
    const arrayBuffer = vi.spyOn(context.req, "arrayBuffer")
    await attachmentHandler()?.(context)

    expect(context.json).toHaveBeenCalledWith(expect.anything(), 401)
    expect(arrayBuffer).not.toHaveBeenCalled()
    expect(fetchSpy.mock.calls.every(([url]) => !String(url).includes("/attachments"))).toBe(true)
  })

  it("does not read the body of a request aimed at a message that is not a draft", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(graph({ id: "sent", isDraft: false })))
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    const context = request({ messageId: "sent", name: "report.pdf", encoding: "base64" })
    const arrayBuffer = vi.spyOn(context.req, "arrayBuffer")
    await attachmentHandler()?.(context)

    expect(context.json).toHaveBeenCalledWith(expect.anything(), 409)
    expect(arrayBuffer).not.toHaveBeenCalled()
  })

  it("refuses a declared length over the 150 MB limit without reading the body", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(graph({ id: "draft", isDraft: true })))
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    const context = request({ messageId: "draft", name: "big.bin", encoding: "base64" })
    context.req.header = (name: string) =>
      name.toLowerCase() === "content-length" ? String(300 * 1024 * 1024) : "Bearer TOKEN"
    const arrayBuffer = vi.spyOn(context.req, "arrayBuffer")
    await attachmentHandler()?.(context)

    expect(context.json).toHaveBeenCalledWith(expect.anything(), 413)
    expect(arrayBuffer).not.toHaveBeenCalled()
  })

  it("passes an apiVersion override through to the attachment request", async () => {
    const fetchSpy = vi
      .fn()
      .mockResolvedValueOnce(graph({ id: "draft", isDraft: true }))
      .mockResolvedValueOnce(graph({ id: "att" }, 201))
    vi.stubGlobal("fetch", fetchSpy)
    await import("../src/index")
    await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

    await attachmentHandler()?.(
      request({ messageId: "draft", name: "report.pdf", encoding: "base64", apiVersion: "beta" }),
    )

    expect(fetchSpy.mock.calls.map(([url]) => String(url).split("/me/")[0])).toEqual([
      "https://graph.microsoft.com/beta",
      "https://graph.microsoft.com/beta",
    ])
  })

  it.each(["/upload", "/upload/attachment"])(
    "rejects an unknown apiVersion on %s before calling Graph",
    async (path) => {
      const fetchSpy = vi.fn()
      vi.stubGlobal("fetch", fetchSpy)
      await import("../src/index")
      await vi.waitFor(() => expect(server.start).toHaveBeenCalled())

      const context = request({
        path: "/me/drive/root:/x.txt:/content",
        messageId: "draft",
        name: "report.pdf",
        apiVersion: "v1.0/../../users",
      })
      await server.post.mock.calls.find(([mounted]) => mounted === path)?.[1](context)

      expect(context.json).toHaveBeenCalledWith({ error: 'apiVersion must be "v1.0" or "beta"' }, 400)
      expect(fetchSpy).not.toHaveBeenCalled()
    },
  )
})
