import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { afterEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import {
  createPlannerBucket,
  listPlannerBuckets,
  listPlans,
  updatePlannerTask,
  updatePlannerTaskDetails,
} from "../src/tools/planner-tools"

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

  // list_plans must fan out over group memberships — /me/planner/plans alone omits group-owned plans
  // the user wasn't explicitly added to, silently showing an empty board.
  it("aggregates plans across the user's groups and dedupes by id", async () => {
    const fetchMock = vi.fn((url: string) => {
      if (url.includes("/me/memberOf")) return Promise.resolve(response({ value: [{ id: "g1" }, { id: "g2" }] }))
      if (url.includes("/me/planner/plans")) return Promise.resolve(response({ value: [{ id: "p1", title: "Mine" }] }))
      if (url.includes("/groups/g1/planner/plans"))
        return Promise.resolve(
          response({
            value: [
              { id: "p1", title: "Mine" },
              { id: "p2", title: "G1 Plan" },
            ],
          }),
        )
      if (url.includes("/groups/g2/planner/plans"))
        return Promise.resolve(response({ value: [{ id: "p3", title: "G2 Plan" }] }))
      return Promise.resolve(response({ value: [] }))
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const out = (await listPlans()).fold(
      () => "",
      (v) => v,
    )

    expect(out).toContain("G1 Plan")
    expect(out).toContain("G2 Plan")
    expect(out.match(/Mine/g)?.length).toBe(1) // p1 seen in /me AND g1 → deduped to one
  })

  it("skips a failing group instead of blanking the whole board", async () => {
    const fetchMock = vi.fn((url: string) => {
      if (url.includes("/me/memberOf")) return Promise.resolve(response({ value: [{ id: "g1" }, { id: "gBad" }] }))
      if (url.includes("/me/planner/plans")) return Promise.resolve(response({ value: [] }))
      if (url.includes("/groups/g1/planner/plans"))
        return Promise.resolve(response({ value: [{ id: "p2", title: "G1 Plan" }] }))
      if (url.includes("/groups/gBad/planner/plans"))
        return Promise.resolve(response({ error: "forbidden" }, false, 403))
      return Promise.resolve(response({ value: [] }))
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await listPlans()

    expect(result.isRight()).toBe(true)
    expect(result.value as string).toContain("G1 Plan")
  })

  // Details edit/remove: read-modify-write merge + skip-not-found (Graph's null-on-missing-key is a
  // silent no-op, so the tool reports it rather than pretending it worked).
  const detailsSnapshot = {
    "@odata.etag": 'W/"1"',
    checklist: {
      g1: { "@odata.type": "#microsoft.graph.plannerChecklistItem", title: "Alpha", isChecked: false },
    },
    references: {
      "https%3A//x%2Ecom": { "@odata.type": "#microsoft.graph.plannerExternalReference", alias: "X", type: "Other" },
    },
  }

  const capturePatch = (bodyRef: { value: unknown }) =>
    vi.fn((_url: string, init?: RequestInit) => {
      if ((init?.method ?? "GET") === "PATCH") {
        bodyRef.value = JSON.parse(String(init?.body))
        return Promise.resolve(response({}, true, 204))
      }
      return Promise.resolve(response(detailsSnapshot))
    })

  it("update_checklist merges — an omitted field keeps its current value", async () => {
    const body: { value: unknown } = { value: null }
    vi.stubGlobal("fetch", capturePatch(body))

    initializeGraphClient(auth)
    const result = await updatePlannerTaskDetails({ task_id: "t", update_checklist: [{ id: "g1", isChecked: true }] })

    expect(result.isRight()).toBe(true)
    const item = (body.value as { checklist: Record<string, { title: string; isChecked: boolean }> }).checklist.g1
    expect(item.title).toBe("Alpha") // preserved, not blanked
    expect(item.isChecked).toBe(true) // changed
  })

  it("remove_checklist nulls the key; a missing key is reported as skipped, not silently succeeded", async () => {
    const body: { value: unknown } = { value: null }
    vi.stubGlobal("fetch", capturePatch(body))

    initializeGraphClient(auth)
    const result = await updatePlannerTaskDetails({ task_id: "t", remove_checklist: ["g1", "gMissing"] })

    const msg = result.fold(
      () => "",
      (v) => v,
    )
    const checklist = (body.value as { checklist: Record<string, unknown> }).checklist
    expect(checklist.g1).toBeNull()
    expect("gMissing" in checklist).toBe(false)
    expect(msg).toContain("Skipped (not found)")
    expect(msg).toContain("gMissing")
  })

  it("list_planner_buckets formats a plan's buckets", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve(
          response({
            value: [
              { id: "b1", name: "To Do" },
              { id: "b2", name: "Done" },
            ],
          }),
        ),
      ),
    )

    initializeGraphClient(auth)
    const out = (await listPlannerBuckets({ plan_id: "p" })).fold(
      () => "",
      (v) => v,
    )
    expect(out).toContain("To Do")
    expect(out).toContain("b2")
  })

  it("create_planner_bucket posts the name and planId", async () => {
    const body: { value: unknown } = { value: null }
    vi.stubGlobal(
      "fetch",
      vi.fn((_url: string, init?: RequestInit) => {
        if (init?.method === "POST") body.value = JSON.parse(String(init?.body))
        return Promise.resolve(response({ id: "bNew", name: "Backlog" }))
      }),
    )

    initializeGraphClient(auth)
    const result = await createPlannerBucket({ plan_id: "p", name: "Backlog" })

    expect(result.isRight()).toBe(true)
    const posted = body.value as { name: string; planId: string }
    expect(posted.name).toBe("Backlog")
    expect(posted.planId).toBe("p")
  })

  // Item 6: update_planner_task auto-fetches the task ETag when the caller omits it (consistent with
  // update_planner_task_details), but honors a pinned ETag for strict optimistic concurrency.
  it("update_planner_task auto-fetches the ETag when none is provided", async () => {
    const calls: Array<{ method: string; ifMatch?: string }> = []
    const fetchMock = vi.fn((_url: string, init?: RequestInit) => {
      const method = init?.method ?? "GET"
      const ifMatch = (init?.headers as Record<string, string> | undefined)?.["If-Match"]
      calls.push({ method, ifMatch })
      if (method === "GET") return Promise.resolve(response({ "@odata.etag": 'W/"taskE"', id: "t" }))
      return Promise.resolve(response({}, true, 204))
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await updatePlannerTask({ task_id: "t", percent_complete: 100 })

    expect(result.isRight()).toBe(true)
    expect(calls.map((c) => c.method)).toEqual(["GET", "PATCH"])
    expect(calls[1]?.ifMatch).toBe('W/"taskE"')
  })

  it("update_planner_task uses a provided ETag without fetching", async () => {
    const methods: string[] = []
    const fetchMock = vi.fn((_url: string, init?: RequestInit) => {
      methods.push(init?.method ?? "GET")
      return Promise.resolve(response({}, true, 204))
    })
    vi.stubGlobal("fetch", fetchMock)

    initializeGraphClient(auth)
    const result = await updatePlannerTask({ task_id: "t", etag: 'W/"pinned"', title: "X" })

    expect(result.isRight()).toBe(true)
    expect(methods).toEqual(["PATCH"]) // no GET — the caller's ETag is used directly
  })
})
