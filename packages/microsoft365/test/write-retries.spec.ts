// AGENTS.md: a write is retried only on 429. After a 503, 504, timeout or network error it may
// already have run, so a retry could repeat it, or report a delete that landed as "not found".
// These were retried before, because core counted PUT and DELETE as safe to repeat.
import { MAY_HAVE_LANDED } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import { deleteContact } from "../src/tools/contacts-tools"

const gatewayTimeout = () => Promise.resolve(new Response("", { status: 504 }))
const fetchSpy = vi.fn(gatewayTimeout)

beforeEach(() => {
  fetchSpy.mockClear()
  fetchSpy.mockImplementation(gatewayTimeout)
  vi.stubGlobal("fetch", fetchSpy)
  vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "3")
  vi.stubEnv("MS365_GRAPH_BASE_BACKOFF_MS", "0")
  vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
})
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

const client = () => initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("TOKEN")) })

describe("a write that got a 504", () => {
  it.each([
    ["deleteMessage (delete_draft)", () => client().deleteMessage("m1")],
    ["deleteEvent", () => client().deleteEvent("e1")],
    ["deleteContact", () => client().deleteContact("c1")],
    ["deleteOnenotePage", () => client().deleteOnenotePage("p1")],
    ["deleteTodoList", () => client().deleteTodoList("l1")],
    ["deleteTodoTask", () => client().deleteTodoTask("l1", "t1")],
    ["deleteDriveItem", () => client().deleteDriveItem("i1")],
    ["deleteMessageRule", () => client().deleteMessageRule("r1")],
    ["graph_query DELETE", () => client().graphQuery("DELETE", "/me/events/e1")],
    ["graph_query PUT", () => client().graphQuery("PUT", "/me/drive/items/i1/content", { a: 1 })],
  ])("%s is sent once, and says it may have gone through", async (_, call) => {
    const result = await call()
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect(result.isLeft()).toBe(true)
    expect((result.value as { message: string }).message).toContain(MAY_HAVE_LANDED)
  })

  it("reaches the tool's message too", async () => {
    // delete_contact reads the contact first; the 504 is the delete's.
    fetchSpy.mockImplementationOnce(() =>
      Promise.resolve(new Response(JSON.stringify({ id: "c1", displayName: "Ada" }))),
    )
    client()
    const result = await deleteContact({ contact_id: "c1" })
    expect(fetchSpy).toHaveBeenCalledTimes(2)
    expect((result.value as Error).message).toContain(MAY_HAVE_LANDED)
  })

  it("a network error on a write is sent once and says the same", async () => {
    fetchSpy.mockImplementation(() => Promise.reject(new TypeError("fetch failed")))
    const result = await client().deleteEvent("e1")
    expect(fetchSpy).toHaveBeenCalledTimes(1)
    expect((result.value as { message: string }).message).toContain(MAY_HAVE_LANDED)
  })

  it("a write that got a 429 is retried, and its success reported", async () => {
    fetchSpy
      .mockImplementationOnce(() => Promise.resolve(new Response("", { status: 429, headers: { "Retry-After": "0" } })))
      .mockImplementationOnce(() => Promise.resolve(new Response(null, { status: 204 })))
    const result = await client().deleteContact("c1")
    expect(result.isRight()).toBe(true)
    expect(fetchSpy).toHaveBeenCalledTimes(2)
  })
})

describe("a read that got a 504", () => {
  it("is retried, and its error carries no note about a change", async () => {
    const result = await client().getContact("c1")
    expect(fetchSpy).toHaveBeenCalledTimes(4)
    expect((result.value as { message: string }).message).not.toContain(MAY_HAVE_LANDED)
  })
})
