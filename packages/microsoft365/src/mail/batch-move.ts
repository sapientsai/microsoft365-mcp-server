// Moving thousands of messages one request at a time is the bottleneck of a mailbox
// sweep: Graph has no bulk move, but it does have JSON batching — up to 20 sub-requests
// per round-trip, each judged individually. This module builds those batches and reads
// their results. It is pure apart from the injected `send` and `wait`, so the retry
// logic can be tested without a mailbox or a clock.

import type { Either } from "functype/either"

import type { GraphApiError, GraphBatchResponse } from "../types"

/** Graph's hard limit on sub-requests per $batch call. */
export const BATCH_SIZE = 20

/** How many times a throttled (429) or transient (503/504) sub-request is retried. */
const MAX_ATTEMPTS = 3

export const chunk = <T>(items: ReadonlyArray<T>, size: number): ReadonlyArray<ReadonlyArray<T>> =>
  items.length === 0 ? [] : [items.slice(0, size), ...chunk(items.slice(size), size)]

export type BatchMoveRequest = {
  readonly id: string
  readonly method: "POST"
  readonly url: string
  readonly headers: { readonly "Content-Type": "application/json" }
  readonly body: { readonly destinationId: string }
}

// Sub-request ids are positions ("0".."19"), not the message ids. Graph compares
// batch ids case-insensitively, and message ids are base64 — two in the same batch
// that differ only by letter case were rejected with "Request Id ... has to be unique
// in a batch" (11% of a 14,000-message run). The response is matched back by that
// positional id, never by array order, so reordering by Graph is still safe.
export const buildMoveBatch = (
  messageIds: ReadonlyArray<string>,
  prefix: string,
  destinationId: string,
): ReadonlyArray<BatchMoveRequest> =>
  messageIds.map((id, index) => ({
    id: String(index),
    method: "POST",
    url: `${prefix}/messages/${id}/move`,
    headers: { "Content-Type": "application/json" },
    body: { destinationId },
  }))

export type MoveOutcome = {
  readonly id: string
  readonly status: number
  readonly error?: string
  readonly retryAfterMs?: number
}

const RETRYABLE = new Set([429, 503, 504])

const headerValue = (headers: Record<string, string> | undefined, name: string): string | undefined => {
  if (!headers) return undefined
  const match = Object.entries(headers).find(([k]) => k.toLowerCase() === name.toLowerCase())
  return match?.[1]
}

export const parseBatchResponses = (
  requested: ReadonlyArray<string>,
  response: GraphBatchResponse,
): ReadonlyArray<MoveOutcome> => {
  const byId = new Map(response.responses.map((r) => [r.id, r]))
  return requested.map((id, index) => {
    const r = byId.get(String(index))
    // A sub-request Graph did not answer at all is treated as failed, not moved:
    // "unknown" must never be reported as success.
    if (r === undefined) return { id, status: 0, error: "No response for this message in the batch" }
    if (r.status >= 200 && r.status < 300) return { id, status: r.status }
    const body = r.body as { error?: { code?: string; message?: string } } | undefined
    const error = body?.error ? `${body.error.code ?? ""} ${body.error.message ?? ""}`.trim() : `HTTP ${r.status}`
    const retryAfter = headerValue(r.headers, "Retry-After")
    const retryAfterMs = retryAfter !== undefined && /^\d+$/.test(retryAfter) ? Number(retryAfter) * 1000 : undefined
    return { id, status: r.status, error, retryAfterMs }
  })
}

export type BatchSender = (
  requests: ReadonlyArray<BatchMoveRequest>,
) => Promise<Either<GraphApiError, GraphBatchResponse>>

export type MoveRunResult = {
  readonly moved: ReadonlyArray<string>
  readonly failed: ReadonlyArray<MoveOutcome>
}

const isRetryable = (o: MoveOutcome): boolean => RETRYABLE.has(o.status)

// One chunk, with retries for the sub-requests Graph throttled. Retries only the
// throttled ids: re-sending a whole chunk would move already-moved messages again
// (a 404 on the old id), which reads as failure.
const runChunk = async (
  ids: ReadonlyArray<string>,
  prefix: string,
  destinationId: string,
  send: BatchSender,
  wait: (ms: number) => Promise<void>,
  attempt: number = 1,
): Promise<MoveRunResult> => {
  const result = await send(buildMoveBatch(ids, prefix, destinationId))
  if (result.isLeft()) {
    const { message } = result.value as GraphApiError
    return { moved: [], failed: ids.map((id) => ({ id, status: 0, error: message })) }
  }

  const outcomes = parseBatchResponses(ids, result.value as GraphBatchResponse)
  const moved = outcomes.filter((o) => o.error === undefined).map((o) => o.id)
  const failed = outcomes.filter((o) => o.error !== undefined)
  const retryable = failed.filter(isRetryable)

  if (retryable.length === 0 || attempt >= MAX_ATTEMPTS) return { moved, failed }

  // Honour the longest Retry-After in the chunk; when Graph gave none, back off for
  // longer each attempt so a persistently throttled mailbox is not hammered.
  const advised = retryable.map((o) => o.retryAfterMs).filter((ms): ms is number => ms !== undefined)
  const delay = advised.length > 0 ? Math.max(...advised) : attempt * 2000
  await wait(delay)

  const retried = await runChunk(
    retryable.map((o) => o.id),
    prefix,
    destinationId,
    send,
    wait,
    attempt + 1,
  )
  return {
    moved: [...moved, ...retried.moved],
    failed: [...failed.filter((o) => !isRetryable(o)), ...retried.failed],
  }
}

/**
 * Move every id, in chunks of BATCH_SIZE, sequentially. Sequential on purpose: Graph
 * throttles per mailbox, and parallel batches only convert throughput into 429s.
 */
export const runMoveBatches = async (
  messageIds: ReadonlyArray<string>,
  prefix: string,
  destinationId: string,
  send: BatchSender,
  wait: (ms: number) => Promise<void> = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  onProgress?: (done: number, total: number, failed: number, chunkResult: MoveRunResult) => void,
  // Optional pause between batches. Exchange allows roughly 10,000 requests per mailbox
  // per ten minutes; a pause of ~750 ms keeps 20-per-batch under that so a long run
  // never trips the throttle instead of relying on Retry-After to recover from it.
  paceMs: number = 0,
): Promise<MoveRunResult> => {
  const chunks = chunk(messageIds, BATCH_SIZE)
  return chunks.reduce<Promise<MoveRunResult>>(
    async (acc, ids, index) => {
      const soFar = await acc
      if (index > 0 && paceMs > 0) await wait(paceMs)
      const result = await runChunk(ids, prefix, destinationId, send, wait)
      const next = { moved: [...soFar.moved, ...result.moved], failed: [...soFar.failed, ...result.failed] }
      onProgress?.(next.moved.length + next.failed.length, messageIds.length, next.failed.length, result)
      return next
    },
    Promise.resolve({ moved: [], failed: [] }),
  )
}
