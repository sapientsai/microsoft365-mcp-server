// Every tool, every string parameter: a hostile value must never change which Graph resource a
// request reaches. fetch normalises a URL before sending (dot segments, backslashes, tabs, "#"),
// so an ID that went into a path as written could point a tool somewhere else entirely. This runs
// each tool once with benign values, then once per string parameter with each hostile value, and
// compares what reached fetch. New tools are covered without anyone remembering to add them.
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import { Right } from "functype/either"
import type { SomaServerInstance } from "somamcp/node"
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest"

const captured = vi.hoisted(() => ({ server: undefined as SomaServerInstance | undefined }))

vi.mock("somamcp/node", async (original) => {
  const real = await original<typeof import("somamcp/node")>()
  return {
    ...real,
    createServer: (options: Parameters<typeof real.createServer>[0]) => {
      const server = real.createServer(options)
      captured.server = server
      return { ...server, start: async () => undefined }
    },
  }
})
vi.mock("../src/auth", async (original) => ({
  ...(await original<typeof import("../src/auth")>()),
  initializeAuth: vi.fn(async () => Right(undefined)),
  getAccessToken: vi.fn(async () => Right("TOKEN")),
}))

type Sent = { readonly method: string; readonly url: string }
type JsonSchema = {
  type?: string | string[]
  enum?: unknown[]
  items?: JsonSchema
  properties?: Record<string, JsonSchema>
  required?: string[]
  anyOf?: JsonSchema[]
  minimum?: number
}
type Tool = { name: string; inputSchema: JsonSchema }

const BASE = "http://localhost:3000"
const sent: Sent[] = []

// One answer for every Graph call: enough shape for most tools to carry on to their next request.
const graphAnswer = () =>
  new Response(
    JSON.stringify({ value: [], id: "BENIGNID", "@odata.etag": 'W/"1"', isDraft: true, name: "a.txt", size: 1 }),
    { status: 200, headers: { "Content-Type": "application/json" } },
  )

// Tools that touch the local disk with a path parameter are left out: their paths are local, not
// Graph's, and the hostile values would be read or written as files.
const SKIPPED_TOOLS = new Set(["save_attachment", "upload_file_from_path", "set_access_token", "switch_account"])

// Parameters that are paths by design: they may hold "/" and so add segments. They must still go
// out exactly as written, which the normalisation check below covers.
const PATH_PARAMS = new Set(["path", "folder_path"])

const HOSTILE = ["MARK/forward#", "../../users/v@x.com/MARK", "MARK?x=1#y", ".\t.", "MARK\\..\\x"]

// Values that pass each tool's own checks, chosen from the parameter's name, word by word.
const benignString = (tool: string, name: string): string => {
  const words = new Set(name.toLowerCase().split(/_|(?=[A-Z])/))
  const has = (...candidates: string[]) => candidates.some((word) => words.has(word))
  if (name === "path") {
    if (tool === "graph_query") return "/me/BENIGN"
    if (tool === "read_document") return "/me/drive/items/BENIGN"
    return "/me/drive/root:/BENIGN.txt:/content"
  }
  if (tool === "graph_query" && name === "body") return "{}"
  if (name === "method") return "GET"
  if (has("mailbox", "email", "address", "to", "recipient", "upn")) return "benign@example.com"
  if (has("zone")) return "UTC"
  if (has("birthday")) return "2026-01-01"
  if (has("end")) return "2026-01-02T09:00:00"
  if (has("date", "time", "start", "due")) return "2026-01-01T09:00:00"
  if (has("url")) return "https://teams.microsoft.com/l/meetup-join/BENIGN"
  if (has("content") && has("type")) return "text/plain"
  if (has("color")) return "red"
  if (has("extension")) return "com.example.benign"
  return "BENIGN"
}

const typeOf = (schema: JsonSchema): string | undefined => {
  if (schema.anyOf) return schema.anyOf.map(typeOf).find((type) => type !== undefined && type !== "null")
  return Array.isArray(schema.type) ? schema.type.find((type) => type !== "null") : schema.type
}

const benignValue = (tool: string, name: string, schema: JsonSchema): unknown => {
  if (schema.enum) return schema.enum[0]
  switch (typeOf(schema)) {
    case "string":
      return benignString(tool, name)
    case "number":
    case "integer":
      return Math.max(schema.minimum ?? 1, 1)
    case "boolean":
      return false
    case "array":
      return [benignValue(tool, name, schema.items ?? {})]
    case "object":
      return {}
    default:
      return undefined
  }
}

const isFreeString = (schema: JsonSchema): boolean => typeOf(schema) === "string" && schema.enum === undefined

// Only the required parameters: an optional one that fails its own check would stop the call.
const requiredArgs = (tool: Tool): Record<string, unknown> =>
  Object.fromEntries(
    Object.entries(tool.inputSchema.properties ?? {})
      .filter(([name]) => tool.inputSchema.required?.includes(name))
      .map(([name, schema]) => [name, benignValue(tool.name, name, schema)] as const)
      .filter(([, value]) => value !== undefined),
  )

const rpc = async (server: SomaServerInstance, method: string, params: unknown) => {
  const response = await server.fetch(
    new Request(`${BASE}/mcp`, {
      method: "POST",
      headers: {
        accept: "application/json, text/event-stream",
        "content-type": "application/json",
        "mcp-protocol-version": "2025-06-18",
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method, params }),
    }),
  )
  const text = await response.text()
  const json = text.startsWith("{") ? text : (text.split("\n").find((line) => line.startsWith("data: ")) ?? "").slice(6)
  return JSON.parse(json).result
}

const run = async (server: SomaServerInstance, name: string, args: Record<string, unknown>): Promise<Sent[]> => {
  sent.length = 0
  await rpc(server, "tools/call", { name, arguments: args })
  return [...sent]
}

const rawPathname = (url: string): string => url.replace(/^https?:\/\/[^/]+/, "").split("?")[0]!
const segments = (url: string): number => new URL(url).pathname.split("/").length
const queryKeys = (url: string): string => [...new URL(url).searchParams.keys()].sort().join("&")

// What fetch sends is the URL after parsing. It must be the URL that was written: same path, no
// fragment, nothing resolved, dropped or rewritten on the way.
const problemsIn = (request: Sent): string[] => {
  const parsed = new URL(request.url)
  return [
    request.url.includes("#") || parsed.hash !== "" ? "has a fragment" : undefined,
    rawPathname(request.url) !== parsed.pathname ? `path rewritten to ${parsed.pathname}` : undefined,
  ].filter((problem): problem is string => problem !== undefined)
}

describe("caller strings never change the Graph resource a tool reaches", () => {
  let server: SomaServerInstance
  let tools: Tool[] = []
  let storeDirectory = ""

  beforeAll(async () => {
    storeDirectory = mkdtempSync(join(tmpdir(), "ms365-path-property-"))
    vi.stubEnv("MS365_AUTH_MODE", "client-token")
    vi.stubEnv("TRANSPORT_TYPE", "httpStream")
    vi.stubEnv("MS365_ACCESS_TOKEN", "a.b.c")
    vi.stubEnv("TOKEN_STORAGE_PATH", storeDirectory)
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
    vi.stubGlobal(
      "fetch",
      vi.fn((input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url
        sent.push({ method: (init?.method ?? "GET").toUpperCase(), url })
        return Promise.resolve(graphAnswer())
      }),
    )
    await import("../src/index")
    await vi.waitFor(() => expect(captured.server).toBeDefined())
    server = captured.server as SomaServerInstance
    tools = ((await rpc(server, "tools/list", {})) as { tools: Tool[] }).tools.filter(
      (tool) => !SKIPPED_TOOLS.has(tool.name),
    )
  })
  afterAll(() => {
    vi.unstubAllGlobals()
    vi.unstubAllEnvs()
    rmSync(storeDirectory, { recursive: true, force: true })
  })

  it("covers the tool list", () => {
    expect(tools.length).toBeGreaterThan(70)
  })

  it("keeps every request where the benign call sent it", { timeout: 120_000 }, async () => {
    const failures: string[] = []
    // eslint-disable-next-line functype/no-let -- counts what the run actually exercised
    let compared = 0
    const exercised = new Set<string>()
    for (const tool of tools) {
      const base = requiredArgs(tool)
      const params = Object.entries(tool.inputSchema.properties ?? {}).filter(
        ([, schema]) => isFreeString(schema) || (typeOf(schema) === "array" && isFreeString(schema.items ?? {})),
      )
      for (const [param, schema] of params) {
        const wrap = (text: string) => (typeOf(schema) === "array" ? [text] : text)
        // The baseline sends the same parameter with a benign value, so the only difference is the value.
        const expected = await run(server, tool.name, { ...base, [param]: wrap(benignString(tool.name, param)) })
        if (expected.length > 0) exercised.add(tool.name)
        expected.flatMap(problemsIn).forEach((problem) => failures.push(`${tool.name}.${param} (benign): ${problem}`))
        for (const hostile of HOSTILE) {
          const actual = await run(server, tool.name, { ...base, [param]: wrap(hostile) })
          const where = `${tool.name}.${param}=${JSON.stringify(hostile)}`
          actual.forEach((request, index) => {
            problemsIn(request).forEach((problem) => failures.push(`${where}: ${problem} (${request.url})`))
            const before = expected[index]
            if (before === undefined) return
            compared += 1
            if (request.method !== before.method)
              failures.push(`${where}: method ${request.method}, not ${before.method}`)
            if (PATH_PARAMS.has(param)) return
            if (segments(request.url) !== segments(before.url)) {
              failures.push(
                `${where}: reached ${new URL(request.url).pathname}, not the shape of ${new URL(before.url).pathname}`,
              )
            }
            if (queryKeys(request.url) !== queryKeys(before.url)) {
              failures.push(`${where}: query became ${new URL(request.url).search}`)
            }
          })
        }
      }
    }
    if (process.env.PROBE)
      console.log(
        "UNEXERCISED",
        tools
          .filter((t) => !exercised.has(t.name))
          .map((t) => t.name)
          .join(" "),
      )
    expect(failures).toEqual([])
    // A run that compared nothing would pass without testing anything.
    expect(compared).toBeGreaterThan(200)
    expect(exercised.size).toBeGreaterThan(60)
  })
})
