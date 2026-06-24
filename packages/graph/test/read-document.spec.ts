import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Left, Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { buildReadDocumentTool } from "../src/tools/read-document"

const auth = (token = "TOK"): AuthStrategy => ({ getAccessToken: () => Promise.resolve(Right(token)) })

const textResponse = (body: string, contentType = "text/plain", ok = true, status = 200) =>
  ({
    ok,
    status,
    statusText: ok ? "OK" : "Error",
    headers: new Headers({ "content-type": contentType }),
    arrayBuffer: () => Promise.resolve(Buffer.from(body)),
    json: () => Promise.resolve(JSON.parse(body)),
  }) as Response

describe("read_document tool", () => {
  it("downloads with the injected Bearer and returns extracted text", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(textResponse("Hello world", "text/plain")))
    const tool = buildReadDocumentTool(auth("INJECTED"), fetchImpl as unknown as typeof fetch)

    const out = await tool.execute({ path: "/me/drive/items/1/content", api_version: "v1.0", max_chars: 50000 })

    const [url, init] = fetchImpl.mock.calls[0] as [string, RequestInit]
    expect(url).toBe("https://graph.microsoft.com/v1.0/me/drive/items/1/content")
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer INJECTED")
    expect(out).toContain("Hello world")
  })

  it("appends a conversion format as a query param", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(textResponse("x")))
    await buildReadDocumentTool(auth(), fetchImpl as unknown as typeof fetch).execute({
      path: "/me/drive/items/1/content",
      api_version: "v1.0",
      format: "pdf",
      max_chars: 50000,
    })
    expect((fetchImpl.mock.calls[0] as [string])[0]).toBe(
      "https://graph.microsoft.com/v1.0/me/drive/items/1/content?format=pdf",
    )
  })

  it("truncates content beyond max_chars", async () => {
    const big = "a".repeat(5000)
    const fetchImpl = vi.fn(() => Promise.resolve(textResponse(big)))
    const out = await buildReadDocumentTool(auth(), fetchImpl as unknown as typeof fetch).execute({
      path: "/me/drive/items/1/content",
      api_version: "v1.0",
      max_chars: 1000,
    })
    expect(out).toContain("[truncated at 1,000 chars")
    expect(out).toContain("full document is 5,000 chars")
  })

  it("throws the Graph error message on a JSON error response", async () => {
    const fetchImpl = vi.fn(() =>
      Promise.resolve(
        textResponse(JSON.stringify({ error: { message: "Item not found" } }), "application/json", false, 404),
      ),
    )
    await expect(
      buildReadDocumentTool(auth(), fetchImpl as unknown as typeof fetch).execute({
        path: "/me/drive/items/x/content",
        api_version: "v1.0",
        max_chars: 50000,
      }),
    ).rejects.toThrow("Item not found")
  })

  it("throws when the auth strategy fails (no download)", async () => {
    const fetchImpl = vi.fn()
    const failing: AuthStrategy = {
      getAccessToken: () => Promise.resolve(Left({ type: "token", message: "no token" })),
    }
    await expect(
      buildReadDocumentTool(failing, fetchImpl as unknown as typeof fetch).execute({
        path: "/me/drive/items/1/content",
        api_version: "v1.0",
        max_chars: 50000,
      }),
    ).rejects.toThrow("no token")
    expect(fetchImpl).not.toHaveBeenCalled()
  })
})
