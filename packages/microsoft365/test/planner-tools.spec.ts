import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import { updatePlannerTaskDetails } from "../src/tools/planner-tools"

// The task-details object carries its own ETag; a concurrent edit between our read and PATCH yields a
// 412, which updatePlannerTaskDetails absorbs by re-reading the ETag and retrying exactly once.
describe("updatePlannerTaskDetails If-Match retry", () => {
  afterEach(() => vi.unstubAllGlobals())

  const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("t")) }

  const response = (body: unknown, ok = true, status = 200) =>
    ({
      ok,
      status,
      text: () => Promise.resolve(JSON.stringify(body)),
      json: () => Promise.resolve(body),
      headers: new Headers(),
    }) as Response

  const detailsRead = response({ "@odata.etag": 'W/"1"' })

  it("re-reads the ETag and retries once when the first PATCH returns 412", async () => {
    let patchCalls = 0
    const fetchMock = vi.fn((_url: string, init?: RequestInit) => {
      if ((init?.method ?? "GET") === "PATCH") {
        patchCalls += 1
        return Promise.resolve(patchCalls === 1 ? response({ error: "stale" }, false, 412) : response({}, true, 204))
      }
      return Promise.resolve(detailsRead)
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await updatePlannerTaskDetails({ task_id: "abc", description: "hi" })

    expect(result.isRight()).toBe(true)
    // GET → PATCH(412) → GET → PATCH(204)
    expect(fetchMock).toHaveBeenCalledTimes(4)
    expect(patchCalls).toBe(2)
  })

  it("does not retry on a non-412 error", async () => {
    const fetchMock = vi.fn((_url: string, init?: RequestInit) =>
      Promise.resolve((init?.method ?? "GET") === "PATCH" ? response({ error: "forbidden" }, false, 403) : detailsRead),
    )
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await updatePlannerTaskDetails({ task_id: "abc", description: "hi" })

    expect(result.isLeft()).toBe(true)
    // GET → PATCH(403), no retry
    expect(fetchMock).toHaveBeenCalledTimes(2)
  })

  // Regression: encodeURIComponent escaped the slashes too (// → %2F%2F), so Planner rejected the key
  // with "The Authority/Host could not be parsed". The key must keep slashes intact and escape only
  // the property-name-illegal chars. Verified live against Graph 2026-07-12.
  it("encodes reference keys with slashes intact, escaping only : and .", async () => {
    let patchBody: unknown
    const fetchMock = vi.fn((_url: string, init?: RequestInit) => {
      if ((init?.method ?? "GET") === "PATCH") {
        patchBody = JSON.parse(String(init?.body))
        return Promise.resolve(response({}, true, 204))
      }
      return Promise.resolve(detailsRead)
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await updatePlannerTaskDetails({
      task_id: "abc",
      add_references: [{ url: "https://docs.github.com/en/rest" }],
    })

    expect(result.isRight()).toBe(true)
    const keys = Object.keys((patchBody as { references: Record<string, unknown> }).references)
    expect(keys).toEqual(["https%3A//docs%2Egithub%2Ecom/en/rest"])
    expect(keys[0]).not.toContain("%2F") // slashes must NOT be encoded
  })
})
