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
      { id: "b", status: 404, error: "ErrorItemNotFound gone", retryAfterMs: undefined },
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

  it("does not retry a non-transient failure", async () => {
    const send = vi.fn().mockResolvedValue(
      Right({ responses: [{ id: "0", status: 404, body: { error: { code: "ErrorItemNotFound" } } }] }),
    )
    const result = await runMoveBatches(["a"], "/me", "archive", send)
    expect(send).toHaveBeenCalledTimes(1)
    expect(result.failed[0]!.error).toContain("ErrorItemNotFound")
  })

  it("fails the whole chunk when the batch call itself fails", async () => {
    const send = vi.fn().mockResolvedValue(Left({ type: "network", message: "offline" }))
    const result = await runMoveBatches(["a", "b"], "/me", "archive", send)
    expect(result.moved).toEqual([])
    expect(result.failed.map((f) => f.error)).toEqual(["offline", "offline"])
  })
})

describe("runMoveBatches pacing and progress", () => {
  it("pauses between batches by paceMs but not before the first", async () => {
    const ids = Array.from({ length: 45 }, (_, i) => `m${i}`)
    const send = vi.fn(async (requests: ReadonlyArray<{ id: string }>) => Right(ok(requests.map((r) => r.id))))
    const wait = vi.fn(async () => undefined)
    await runMoveBatches(ids, "/me", "archive", send as never, wait, undefined, 750)
    expect(wait.mock.calls).toEqual([[750], [750]])
  })

  it("reports cumulative progress and the chunk's own result after every batch", async () => {
    const ids = Array.from({ length: 25 }, (_, i) => `m${i}`)
    const send = vi.fn(async (requests: ReadonlyArray<{ id: string }>) => Right(ok(requests.map((r) => r.id))))
    const progress = vi.fn()
    await runMoveBatches(ids, "/me", "archive", send as never, undefined, progress)
    expect(progress).toHaveBeenCalledTimes(2)
    expect(progress.mock.calls[0]!.slice(0, 3)).toEqual([20, 25, 0])
    expect(progress.mock.calls[1]!.slice(0, 3)).toEqual([25, 25, 0])
    expect(progress.mock.calls[1]![3].moved).toHaveLength(5)
  })
})
