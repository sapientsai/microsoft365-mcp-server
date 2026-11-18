import { Left, Right } from "functype/either"
import { describe, expect, it, vi } from "vitest"

import { BATCH_SIZE, buildMoveBatch, chunk, parseBatchResponses, runMoveBatches } from "../src/mail/batch-move"
import type { GraphBatchResponse } from "../src/types"

// Responses carry positional ids, matching buildMoveBatch.
const ok = (ids: ReadonlyArray<string>): GraphBatchResponse => ({
  responses: ids.map((_, index) => ({ id: String(index), status: 201 })),
})

describe("chunk", () => {
  it("splits into groups of the given size, last group short", () => {
    expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]])
    expect(chunk([], 2)).toEqual([])
  })
})

describe("buildMoveBatch", () => {
  // Graph compares batch ids case-insensitively; base64 message ids collided.
  it("addresses each move under the mailbox prefix and keys it by position, not message id", () => {
    const [req] = buildMoveBatch(["m1"], "/users/bel%40example.com", "deleteditems")
    expect(req).toEqual({
      id: "0",
      method: "POST",
      url: "/users/bel%40example.com/messages/m1/move",
      headers: { "Content-Type": "application/json" },
      body: { destinationId: "deleteditems" },
    })
  })
})

describe("parseBatchResponses", () => {
  it("reads success and failure per sub-request", () => {
    const outcomes = parseBatchResponses(["a", "b"], {
      responses: [
        { id: "1", status: 404, body: { error: { code: "ErrorItemNotFound", message: "gone" } } },
        { id: "0", status: 201 },
      ],
    })
    expect(outcomes).toEqual([
      { id: "a", status: 201 },
      { id: "b", status: 404, error: "gone (HTTP 404, ErrorItemNotFound)", retryAfterMs: undefined },
    ])
  })

  it("names the status and code when Graph sends an empty message, or no error body at all", () => {
    const outcomes = parseBatchResponses(["a", "b"], {
      responses: [
        { id: "0", status: 503, body: { error: { code: "UnknownError", message: "" } } },
        { id: "1", status: 503 },
      ],
    })
    expect(outcomes.map((o) => o.error)).toEqual([
      "Microsoft Graph API error (HTTP 503, UnknownError)",
      "Microsoft Graph API error (HTTP 503)",
    ])
  })

  // "Unknown" must never be reported as moved.
  it("treats a missing response as a failure", () => {
    const [outcome] = parseBatchResponses(["a"], { responses: [] })
    expect(outcome!.error).toContain("No response")
  })

  it("reads Retry-After on a throttled sub-request", () => {
    const [outcome] = parseBatchResponses(["a"], {
      responses: [{ id: "0", status: 429, headers: { "retry-after": "3" }, body: { error: { code: "TooMany" } } }],
    })
    expect(outcome).toMatchObject({ status: 429, retryAfterMs: 3000 })
  })

  // A Retry-After of an hour would otherwise stall the tool call for an hour per attempt.
  it("caps Retry-After at 60 seconds", () => {
    const [outcome] = parseBatchResponses(["a"], {
      responses: [{ id: "0", status: 429, headers: { "Retry-After": "3600" } }],
    })
    expect(outcome!.retryAfterMs).toBe(60_000)
  })

  it("reads a Retry-After given as an HTTP-date", () => {
    const at = new Date(Date.now() + 5_000).toUTCString()
    const [outcome] = parseBatchResponses(["a"], {
      responses: [{ id: "0", status: 429, headers: { "Retry-After": at } }],
    })
    expect(outcome!.retryAfterMs).toBeGreaterThan(0)
    expect(outcome!.retryAfterMs).toBeLessThanOrEqual(5_000)
  })
})

describe("runMoveBatches", () => {
  it("sends in chunks of BATCH_SIZE, sequentially", async () => {
    const ids = Array.from({ length: 45 }, (_, i) => `m${i}`)
    const send = vi.fn(async (requests: ReadonlyArray<{ id: string }>) => Right(ok(requests.map((r) => r.id))))
    const result = await runMoveBatches(ids, "/me", "archive", send as never)
    expect(send).toHaveBeenCalledTimes(3)
    expect(send.mock.calls.map((c) => c[0].length)).toEqual([BATCH_SIZE, BATCH_SIZE, 5])
    expect(result.moved).toHaveLength(45)
    expect(result.failed).toEqual([])
  })

  // Only the throttled ids are re-sent; the ones that moved must not be moved again.
  it("retries only throttled sub-requests after waiting Retry-After", async () => {
    const send = vi
      .fn()
      .mockResolvedValueOnce(
        Right({
          responses: [
            { id: "0", status: 201 },
            { id: "1", status: 429, headers: { "Retry-After": "1" } },
          ],
        }),
      )
      .mockResolvedValueOnce(Right(ok(["b"])))
    const wait = vi.fn(async () => undefined)
    const result = await runMoveBatches(["a", "b"], "/me", "archive", send, wait)
    expect(wait).toHaveBeenCalledWith(1000)
    expect(send.mock.calls[1]![0].map((r: { url: string }) => r.url)).toEqual(["/me/messages/b/move"])
    expect(result.moved).toEqual(["a", "b"])
    expect(result.failed).toEqual([])
  })

  it("gives up on a persistently throttled id and reports it", async () => {
    const throttled = Right({ responses: [{ id: "0", status: 429 }] })
    const send = vi.fn().mockResolvedValue(throttled)
    const wait = vi.fn(async () => undefined)
    const result = await runMoveBatches(["a"], "/me", "archive", send, wait)
    expect(send).toHaveBeenCalledTimes(3)
    expect(result.moved).toEqual([])
    expect(result.failed).toHaveLength(1)
    expect(result.failed[0]!.status).toBe(429)
  })

  // A move is a POST: a 503/504 can arrive after Graph already moved the message, and
  // re-sending would 404 on the old id and report a moved message as failed.
  it.each([503, 504])("does not retry a sub-request that returned %i", async (status) => {
    const send = vi.fn().mockResolvedValue(Right({ responses: [{ id: "0", status }] }))
    const wait = vi.fn(async () => undefined)
    const result = await runMoveBatches(["a"], "/me", "archive", send, wait)
    expect(send).toHaveBeenCalledTimes(1)
    expect(wait).not.toHaveBeenCalled()
    expect(result.failed.map((f) => f.status)).toEqual([status])
  })

  it("does not retry a non-transient failure", async () => {
    const send = vi
      .fn()
      .mockResolvedValue(
        Right({ responses: [{ id: "0", status: 404, body: { error: { code: "ErrorItemNotFound" } } }] }),
      )
    const result = await runMoveBatches(["a"], "/me", "archive", send)
    expect(send).toHaveBeenCalledTimes(1)
    expect(result.failed[0]!.error).toContain("ErrorItemNotFound")
  })

  // Mirrors live Graph (checked 2026-10-27): ids "a" and "A" in one batch fail the whole
  // request with "Request Id A has to be unique in a batch". Base64 message ids can
  // differ only by case, so they must never be used as batch ids.
  it("moves message ids that differ only by letter case in the same batch", async () => {
    const send = vi.fn(async (requests: ReadonlyArray<{ id: string }>) => {
      const seen = new Set<string>()
      const dup = requests.find((r) => seen.size === seen.add(r.id.toLowerCase()).size)
      return dup
        ? Left({ type: "api", status: 400, message: `Request Id ${dup.id} has to be unique in a batch.` })
        : Right(ok(requests.map((r) => r.id)))
    })
    const result = await runMoveBatches(["AAMkAGQ=", "AAMkAGq="], "/me", "archive", send as never)
    expect(result.failed).toEqual([])
    expect(result.moved).toEqual(["AAMkAGQ=", "AAMkAGq="])
  })

  it("fails the whole chunk when the batch call itself fails", async () => {
    const send = vi.fn().mockResolvedValue(Left({ type: "network", message: "offline" }))
    const result = await runMoveBatches(["a", "b"], "/me", "archive", send)
    expect(result.moved).toEqual([])
    expect(result.failed.map((f) => f.error)).toEqual(["offline", "offline"])
  })
})
