import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"

// A quoted phrase used to go into $search as-is, so its quotes closed the clause early and Graph
// answered "An identifier was expected at position 0". It must reach Graph backslash-escaped.
describe("searchMessages with a quoted phrase", () => {
  afterEach(() => vi.unstubAllGlobals())

  it("sends the phrase's quotes escaped inside the $search clause", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(new Response(JSON.stringify({ value: [] }))))
    vi.stubGlobal("fetch", fetchSpy)

    const result = await initializeGraphClient({ getAccessToken: async () => Right("TOKEN") }).searchMessages(
      '"quarterly report"',
      { $top: 5 },
    )

    expect(result.isRight()).toBe(true)
    const url = String((fetchSpy.mock.calls[0] as unknown[])[0])
    expect(url).toContain('$search="%5C%22quarterly%20report%5C%22"')
  })
})
