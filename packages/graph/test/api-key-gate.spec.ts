import { mintUploadTicket } from "@sapientsai/ms-graph-core"
import { describe, expect, it } from "vitest"

import { authorizesWithApiKey, presentedApiKey } from "../src/auth/api-key-gate"

describe("authorizesWithApiKey", () => {
  it("accepts the raw api key", () => {
    expect(authorizesWithApiKey("SECRET", "SECRET")).toBe(true)
  })

  it("accepts an opaque upload ticket that resolves to the api key", () => {
    const ticket = mintUploadTicket("SECRET")
    expect(authorizesWithApiKey(ticket, "SECRET")).toBe(true)
  })

  it("rejects a wrong key", () => {
    expect(authorizesWithApiKey("wrong", "SECRET")).toBe(false)
  })

  it("rejects a missing bearer", () => {
    expect(authorizesWithApiKey(undefined, "SECRET")).toBe(false)
  })

  it("rejects an unknown ticket-shaped value", () => {
    expect(authorizesWithApiKey("upl_neverminted", "SECRET")).toBe(false)
  })
})

// http.IncomingMessage — the MCP transport path. Headers are a plain lowercased
// object and `url` is path-relative.
const nodeRequest = (headers: Record<string, string>, url: string) => ({ headers, url })

// Hono Request (c.req.raw) — the protected-route path. WHATWG Headers, absolute url.
const fetchRequest = (headers: Record<string, string>, url: string) => ({
  headers: new Headers(headers),
  url,
})

describe("presentedApiKey", () => {
  it("reads the Authorization bearer on a node request", () => {
    expect(presentedApiKey(nodeRequest({ authorization: "Bearer SECRET" }, "/mcp"))).toBe("SECRET")
  })

  it("reads the Authorization bearer on a fetch request", () => {
    expect(presentedApiKey(fetchRequest({ authorization: "Bearer SECRET" }, "https://h.example/mcp"))).toBe("SECRET")
  })

  // The regression this guards: claude.ai custom connectors carry a URL and no
  // header, so the key can only travel as ?api_key=. Dropping it 401'd every
  // connector configured that way.
  it("falls back to the api_key query parameter on a node request", () => {
    expect(presentedApiKey(nodeRequest({}, "/mcp?api_key=SECRET"))).toBe("SECRET")
  })

  it("falls back to the api_key query parameter on a fetch request", () => {
    expect(presentedApiKey(fetchRequest({}, "https://h.example/mcp?api_key=SECRET"))).toBe("SECRET")
  })

  it("authorizes end to end through the query parameter", () => {
    expect(authorizesWithApiKey(presentedApiKey(nodeRequest({}, "/mcp?api_key=SECRET")), "SECRET")).toBe(true)
  })

  it("accepts an upload ticket presented as a query parameter", () => {
    const ticket = mintUploadTicket("SECRET")
    const url = `/upload?api_key=${encodeURIComponent(ticket)}`
    expect(authorizesWithApiKey(presentedApiKey(nodeRequest({}, url)), "SECRET")).toBe(true)
  })

  it("prefers the header when both are present", () => {
    expect(presentedApiKey(nodeRequest({ authorization: "Bearer FROM_HEADER" }, "/mcp?api_key=FROM_QUERY"))).toBe(
      "FROM_HEADER",
    )
  })

  it("falls through to the query when the header is present but empty", () => {
    expect(presentedApiKey(nodeRequest({ authorization: "Bearer " }, "/mcp?api_key=SECRET"))).toBe("SECRET")
  })

  it("returns undefined when neither is present", () => {
    expect(presentedApiKey(nodeRequest({}, "/mcp"))).toBeUndefined()
  })

  it("returns undefined for an empty api_key parameter", () => {
    expect(presentedApiKey(nodeRequest({}, "/mcp?api_key="))).toBeUndefined()
  })

  it("returns undefined when the request has no recognizable shape", () => {
    expect(presentedApiKey(undefined)).toBeUndefined()
    expect(presentedApiKey({})).toBeUndefined()
    expect(presentedApiKey({ url: 42 })).toBeUndefined()
  })

  // Pinned, not endorsed: a query string decodes "+" to a space, so a key
  // containing "+" cannot travel this way — the same limit the predecessor had.
  // Such keys must use the Authorization header. Percent-encoded "+" is fine.
  it("cannot carry a '+' in a raw query parameter, matching the predecessor", () => {
    expect(presentedApiKey(nodeRequest({}, "/mcp?api_key=a+b"))).toBe("a b")
    expect(presentedApiKey(nodeRequest({}, "/mcp?api_key=a%2Bb"))).toBe("a+b")
  })
})
