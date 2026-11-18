import { describe, expect, it } from "vitest"

import { createGraphRequest } from "../src/graph-request"
import { graphErrorMessage } from "../src/graph-error-message"
import type { AuthStrategy } from "../src/auth-strategy"
import { Right } from "functype/either"

const auth = (): AuthStrategy => ({ getAccessToken: () => Promise.resolve(Right("TOKEN")) }) as AuthStrategy

const unknownError = { error: { code: "UnknownError", message: "" } }

describe("graphErrorMessage", () => {
  it("names the status and code when Graph sends an empty message", () => {
    expect(graphErrorMessage(503, "Service Unavailable", unknownError)).toBe(
      "Microsoft Graph API error (HTTP 503 Service Unavailable, UnknownError)",
    )
  })

  it("treats a whitespace-only message as missing", () => {
    expect(graphErrorMessage(403, "Forbidden", { error: { code: "UnknownError", message: "  \n" } })).toBe(
      "Microsoft Graph API error (HTTP 403 Forbidden, UnknownError)",
    )
  })

  it("keeps Graph's message and adds the status and code after it", () => {
    expect(
      graphErrorMessage(404, "Not Found", { error: { code: "ErrorItemNotFound", message: "The item was not found." } }),
    ).toBe("The item was not found. (HTTP 404 Not Found, ErrorItemNotFound)")
  })

  it("gives the status alone when the body is not a Graph error", () => {
    expect(graphErrorMessage(500, "Internal Server Error", undefined)).toBe(
      "Microsoft Graph API error (HTTP 500 Internal Server Error)",
    )
    expect(graphErrorMessage(502, "Bad Gateway", "<html>gateway</html>")).toBe(
      "Microsoft Graph API error (HTTP 502 Bad Gateway)",
    )
  })

  it("leaves out a blank status text, as HTTP/2 responses have", () => {
    expect(graphErrorMessage(503, "", unknownError)).toBe("Microsoft Graph API error (HTTP 503, UnknownError)")
  })
})

describe("mapHttpError through createGraphRequest", () => {
  // fetchImpl bypasses the resilience wrapper, so the 503 is mapped without spending retries.
  it("reports a 503 UnknownError with an empty message by its status and code", async () => {
    const fetchImpl = () =>
      Promise.resolve(
        new Response(JSON.stringify(unknownError), {
          status: 503,
          statusText: "Service Unavailable",
          headers: { "content-type": "application/json" },
        }),
      )
    const result = await createGraphRequest(auth(), { fetchImpl }).request("GET", "/me/messages")
    expect(result.isLeft()).toBe(true)
    const error = result.value as { type: string; status: number; graphErrorCode?: string; message: string }
    expect(error.type).toBe("api")
    expect(error.status).toBe(503)
    expect(error.graphErrorCode).toBe("UnknownError")
    expect(error.message).toBe("Microsoft Graph API error (HTTP 503 Service Unavailable, UnknownError)")
  })
})
