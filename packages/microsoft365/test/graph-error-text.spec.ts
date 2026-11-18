import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import { listMessages } from "../src/tools/mail-tools"

// End to end through the real client: the tool's error text must say what Graph returned, even
// when Graph's own message is empty (an Exchange outage answers 503 UnknownError with "").
describe("tool error text for a Graph error with an empty message", () => {
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.unstubAllEnvs()
  })

  it("list_messages names the 503 and UnknownError", async () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "0")
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve(
          new Response(JSON.stringify({ error: { code: "UnknownError", message: "" } }), {
            status: 503,
            statusText: "Service Unavailable",
            headers: { "content-type": "application/json" },
          }),
        ),
      ),
    )
    initializeGraphClient({ getAccessToken: async () => Right("TOKEN") })

    const result = await listMessages({})

    expect(result.isLeft()).toBe(true)
    const text = (result.value as Error).message
    expect(text).toContain("Failed to list messages:")
    expect(text).toContain("503")
    expect(text).toContain("UnknownError")
  })
})
