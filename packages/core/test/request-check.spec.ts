import { createServer, type Server } from "node:http"
import type { AddressInfo } from "node:net"

import { Right } from "functype/either"
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from "vitest"

import type { AuthStrategy } from "../src/auth-strategy"
import { createGraphRequest } from "../src/graph-request"
import { isRequestNotSentError, requestProblem, wasRefusedBeforeSending } from "../src/request-check"
import { CircuitBreaker, fetchWithResilience, type ResilienceConfig } from "../src/resilience"
import { uploadDriveContent } from "../src/upload/upload"

const config: ResilienceConfig = {
  maxRetries: 3,
  baseBackoffMs: 0,
  maxBackoffMs: 0,
  fetchTimeoutMs: 2_000,
  circuitFailureThreshold: 1,
  circuitCooldownMs: 30_000,
  circuitDisabled: false,
}
const noSleep = () => Promise.resolve()
const URL_OK = "https://graph.microsoft.com/v1.0/me"

// Inputs fetch refuses to send. Each one used to count as a network failure: up to four per
// call, so two bad calls could open the process-wide breaker. The last field says how this Node
// refuses it: "fetch" throws while building the request, "undici" refuses it at dispatch with an
// error code, and "partial" sends the headers and then fails, so it must be stopped before fetch.
type Refusal = "fetch" | "undici" | "partial"
const UNSENDABLE: ReadonlyArray<readonly [string, string, RequestInit, Refusal]> = [
  ["a line break in a value", "GET", { headers: { Prefer: "a\nb" } }, "fetch"],
  ["a carriage return in a value", "GET", { headers: { Prefer: "a\rb" } }, "fetch"],
  ["a NUL in a value", "GET", { headers: { Prefer: "a\u0000b" } }, "fetch"],
  ["a control character in a value", "GET", { headers: { Prefer: "\u0001" } }, "undici"],
  ["an escape character in a value", "GET", { headers: { Prefer: "a\u001bb" } }, "undici"],
  ["a DEL in a value", "GET", { headers: { Prefer: "a\u007fb" } }, "undici"],
  ["a character above U+00FF in a value", "GET", { headers: { Prefer: "a\u0100b" } }, "fetch"],
  ["a space in a name", "GET", { headers: { "If Match": "x" } }, "fetch"],
  ["an empty name", "GET", { headers: { "": "x" } }, "fetch"],
  ["a non-ASCII name", "GET", { headers: { "Pr\u00e9fer": "x" } }, "fetch"],
  ["an Upgrade header", "GET", { headers: { Upgrade: "h2c" } }, "undici"],
  ["a Keep-Alive header", "GET", { headers: { "Keep-Alive": "5" } }, "undici"],
  ["a Transfer-Encoding header", "POST", { headers: { "Transfer-Encoding": "chunked" }, body: "hello" }, "undici"],
  ["an Expect header", "POST", { headers: { Expect: "100-continue" }, body: "hello" }, "undici"],
  ["a non-numeric Content-Length", "POST", { headers: { "Content-Length": "abc" }, body: "hello" }, "undici"],
  ["a long Content-Length", "POST", { headers: { "Content-Length": "99" }, body: "hello" }, "partial"],
  // "\u00e9" is two bytes. undici sends one, then the server times out or the socket fails.
  ["a Content-Length counting characters", "POST", { headers: { "Content-Length": "1" }, body: "\u00e9" }, "partial"],
  ["a method with a space", " GET", {}, "fetch"],
  ["an empty method", "", {}, "fetch"],
  ["a method with an inner space", "GE T", {}, "fetch"],
  ["CONNECT", "CONNECT", {}, "fetch"],
  ["TRACE", "TRACE", {}, "fetch"],
  ["TRACK", "TRACK", {}, "fetch"],
  ["a GET with a body", "GET", { body: "x" }, "fetch"],
  ["a HEAD with a body", "HEAD", { body: "x" }, "fetch"],
]

describe("requestProblem", () => {
  it.each(UNSENDABLE)("refuses %s", (_, method, init) => {
    expect(requestProblem(URL_OK, { ...init, method })).toBeDefined()
  })

  it.each([
    ["a Latin-1 value", { Prefer: "é" }],
    ["a tab inside a value", { Prefer: "a\tb" }],
    ["a trailing line break, which fetch trims", { Prefer: "a\n" }],
    ["an empty value", { Prefer: "" }],
    ["a TE header", { TE: "trailers" }],
  ])("allows %s", (_, headers) => {
    expect(requestProblem(URL_OK, { method: "GET", headers })).toBeUndefined()
  })

  it("reads a header value that isn't a string as fetch does, as its string form", () => {
    const headers = { "If-Match": undefined } as unknown as Record<string, string>
    expect(requestProblem(URL_OK, { headers })).toBeUndefined()
  })

  it("allows a Content-Length that counts the body's bytes", () => {
    const init = { method: "POST", headers: { "Content-Length": "2" }, body: "é" }
    expect(requestProblem(URL_OK, init)).toBeUndefined()
  })

  it("refuses a URL that doesn't parse", () => {
    expect(requestProblem("not a url", undefined)).toMatch(/not a valid URL/)
  })

  it.each(["ftp://graph.microsoft.com/x", "file:///etc/passwd", "about:blank"])(
    "refuses the non-http URL %s",
    (url) => {
      expect(requestProblem(url, undefined)).toMatch(/not an http\(s\) URL/)
    },
  )

  it("refuses a URL with a user name or password, which fetch won't send", () => {
    expect(requestProblem("https://user:pass@graph.microsoft.com/v1.0/me", undefined)).toMatch(/user name or password/)
  })

  it("caps a long header name in the message", () => {
    const problem = requestProblem(URL_OK, { headers: { [`Upgrade${"x".repeat(5000)}`]: "\u007f" } }) ?? ""
    expect(problem.length).toBeLessThan(200)
  })

  it("reads Headers objects and header pairs as well as records", () => {
    expect(requestProblem(URL_OK, { headers: new Headers({ Upgrade: "h2c" }) })).toMatch(/Upgrade/i)
    expect(requestProblem(URL_OK, { headers: [["Prefer", "a\u007fb"]] })).toMatch(/Prefer/)
  })

  it("never echoes a header value, which may be a token, and keeps the message on one line", () => {
    const problem = requestProblem(URL_OK, { headers: { Authorization: "Bearer secret-token\nInjected: yes" } })
    expect(problem).toBeDefined()
    expect(problem).not.toContain("secret-token")
    expect(problem).not.toContain("\n")
  })

  it("quotes a bad header name so it can't add lines to the message", () => {
    const problem = requestProblem(URL_OK, { headers: { "X\nForged line": "v" } }) ?? ""
    expect(problem.split("\n")).toHaveLength(1)
  })
})

describe("fetchWithResilience with a request fetch can't send", () => {
  it.each(UNSENDABLE)("refuses %s without calling fetch or counting a failure", async (_, method, init) => {
    const fetchImpl = vi.fn()
    const breaker = new CircuitBreaker(1, 30_000, false)
    const call = fetchWithResilience(URL_OK, { ...init, method }, config, breaker, noSleep, fetchImpl)

    await expect(call).rejects.toSatisfy(isRequestNotSentError)
    expect(fetchImpl).not.toHaveBeenCalled()
    expect(breaker.getState().failures).toBe(0)
  })

  it("reports the bad request, not the outage, while the breaker is open", async () => {
    const breaker = new CircuitBreaker(1, 30_000, false)
    breaker.recordFailure()
    const call = fetchWithResilience(URL_OK, { headers: { Prefer: "a\nb" } }, config, breaker, noSleep, vi.fn())
    await expect(call).rejects.toSatisfy(isRequestNotSentError)
  })

  it("lets a healthy call through after many bad ones", async () => {
    const breaker = new CircuitBreaker(5, 30_000, false)
    const bad = Array.from({ length: 20 }, () =>
      fetchWithResilience(URL_OK, { headers: { Prefer: "a\nb" } }, config, breaker, noSleep, vi.fn()).catch(
        () => undefined,
      ),
    )
    await Promise.all(bad)
    const fetchImpl = vi.fn(() => Promise.resolve(new Response("{}")))
    const response = await fetchWithResilience(URL_OK, undefined, config, breaker, noSleep, fetchImpl)
    expect(response.status).toBe(200)
  })

  it("treats undici's refused-to-build codes as not sent: no retry, no failure", async () => {
    const refused = new TypeError("fetch failed", {
      cause: Object.assign(new Error("invalid foo header"), { code: "UND_ERR_INVALID_ARG" }),
    })
    expect(wasRefusedBeforeSending(refused)).toBe(true)
    const fetchImpl = vi.fn(() => Promise.reject(refused))
    const breaker = new CircuitBreaker(1, 30_000, false)

    await expect(fetchWithResilience(URL_OK, undefined, config, breaker, noSleep, fetchImpl)).rejects.toSatisfy(
      isRequestNotSentError,
    )
    expect(fetchImpl).toHaveBeenCalledTimes(1)
    expect(breaker.getState().failures).toBe(0)
  })

  it("still counts a real network failure", async () => {
    const lost = new TypeError("fetch failed", {
      cause: Object.assign(new Error("connect ECONNREFUSED"), { code: "ECONNREFUSED" }),
    })
    expect(wasRefusedBeforeSending(lost)).toBe(false)
    const breaker = new CircuitBreaker(1, 30_000, false)
    const fetchImpl = vi.fn(() => Promise.reject(lost))
    await expect(fetchWithResilience(URL_OK, undefined, config, breaker, noSleep, fetchImpl)).rejects.toBe(lost)
    expect(breaker.getState().failures).toBeGreaterThan(0)
  })
})

// The list above is only right if it matches what this Node's fetch refuses, so the same inputs
// go through the real fetch to a local server: none may reach it, and none may count.
describe("against the real fetch", () => {
  const received: string[] = []
  const server: Server = createServer((req, res) => {
    received.push(`${req.method} ${req.url}`)
    req.resume()
    req.on("end", () => res.end("{}"))
  })
  let base = ""

  beforeAll(async () => {
    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve))
    base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
  })
  afterAll(() => new Promise<void>((resolve) => server.close(() => resolve())))
  afterEach(() => {
    received.length = 0
  })

  // requestProblem refuses every row, so this calls fetch itself: the table is only right if this
  // Node really refuses each input, and the backstop only works if undici's codes are as listed.
  it.each(UNSENDABLE.filter(([, , , how]) => how !== "partial"))(
    "Node's fetch refuses %s without sending it",
    async (_, method, init, how) => {
      const error = await fetch(`${base}/x`, { ...init, method }).then(
        () => undefined,
        (err: unknown) => err,
      )
      expect(error).toBeInstanceOf(TypeError)
      expect(received).toEqual([])
      expect(wasRefusedBeforeSending(error)).toBe(how === "undici")
    },
  )

  it("still sends the values it allows", async () => {
    const breaker = new CircuitBreaker(1, 30_000, false)
    const headers = { Prefer: "é", "If-Match": "a\tb", TE: "trailers" }
    const response = await fetchWithResilience(`${base}/ok`, { headers }, config, breaker, noSleep, fetch)
    expect(response.status).toBe(200)
    expect(received).toEqual(["GET /ok"])
  })
})

describe("createGraphRequest", () => {
  const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("TOK")) }
  afterEach(() => vi.unstubAllGlobals())

  it("answers a request it can't send with an api 400, and sends nothing", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await createGraphRequest(auth).request("GET", "/me", { headers: { Prefer: "a\nb" } })
    expect(result.isLeft()).toBe(true)
    expect(result.value).toMatchObject({ type: "api", status: 400 })
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("does the same for a paginated read", async () => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await createGraphRequest(auth).requestPaginated("/me/messages", { headers: { Upgrade: "h2c" } })
    expect(result.value).toMatchObject({ type: "api", status: 400 })
    expect(fetchSpy).not.toHaveBeenCalled()
  })

  it("keeps its own Authorization whatever the caller passes", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(new Response("{}")))
    vi.stubGlobal("fetch", fetchSpy)
    await createGraphRequest(auth).request("GET", "/me", { headers: { authorization: "Bearer OTHER" } })
    const headers = (fetchSpy.mock.calls[0] as unknown as [string, RequestInit])[1].headers as Record<string, string>
    expect(Object.entries(headers).filter(([name]) => name.toLowerCase() === "authorization")).toEqual([
      ["Authorization", "Bearer TOK"],
    ])
  })

  it("lets a caller's Content-Type replace the default in any case, rather than join it", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(new Response("{}")))
    vi.stubGlobal("fetch", fetchSpy)
    await createGraphRequest(auth).request("POST", "/x", { body: "a", headers: { "content-type": "text/plain" } })
    const headers = new Headers((fetchSpy.mock.calls[0] as unknown as [string, RequestInit])[1].headers)
    expect(headers.get("content-type")).toBe("text/plain")
  })

  it.each([
    ["post", "POST"],
    ["patch", "PATCH"],
  ])("sends a lower-case %s as %s, with its body", async (method, sent) => {
    const fetchSpy = vi.fn(() => Promise.resolve(new Response("{}")))
    vi.stubGlobal("fetch", fetchSpy)
    await createGraphRequest(auth).request(method, "/x", { body: { a: 1 } })
    const init = (fetchSpy.mock.calls[0] as unknown as [string, RequestInit])[1]
    expect(init.method).toBe(sent)
    expect(init.body).toBe('{"a":1}')
  })
})

describe("uploadDriveContent refuses values that can't go in a header", () => {
  const apiBase = "https://graph.microsoft.com/v1.0"
  const path = "/me/drive/root:/a.txt:/content"
  const data = Buffer.from("hello")
  afterEach(() => vi.unstubAllGlobals())

  it.each([
    ["an eTag with a line break", "text/plain", "rename", { ifMatch: '"{A},1"\nX: y' }],
    ["a content type with a line break", "text/plain\nX: y", "rename", {}],
    ["a blank content type", " ", "rename", {}],
    ["an unknown conflict behavior", "text/plain", "rename&x=1", {}],
  ])("refuses %s with a 400 and sends nothing", async (_, contentType, conflict, options) => {
    const fetchSpy = vi.fn()
    vi.stubGlobal("fetch", fetchSpy)
    const result = await uploadDriveContent(apiBase, path, "token", data, contentType, conflict, options)
    expect(result.value).toMatchObject({ type: "api", status: 400 })
    expect(fetchSpy).not.toHaveBeenCalled()
  })
})
