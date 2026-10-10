import { beforeEach, describe, expect, it, vi } from "vitest"

import {
  backoffDelayMs,
  CircuitBreaker,
  CircuitOpenError,
  fetchWithResilience,
  isCircuitOpenError,
  isMethodIdempotent,
  loadResilienceConfig,
  parseRetryAfterMs,
  type ResilienceConfig,
} from "../src/resilience"

const config = (overrides: Partial<ResilienceConfig> = {}): ResilienceConfig => ({
  maxRetries: 3,
  baseBackoffMs: 10,
  maxBackoffMs: 50,
  fetchTimeoutMs: 1_000,
  circuitFailureThreshold: 5,
  circuitCooldownMs: 30_000,
  circuitDisabled: true,
  ...overrides,
})

// A breaker that never opens, so retry tests exercise retry alone.
const openBreaker = () => new CircuitBreaker(5, 30_000, true)

const res = (status: number, headers: Record<string, string> = {}) =>
  ({
    ok: status >= 200 && status < 300,
    status,
    statusText: "",
    headers: new Headers(headers),
    arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)),
    text: () => Promise.resolve(""),
    json: () => Promise.resolve({}),
  }) as Response

const noSleep = () => Promise.resolve()

describe("parseRetryAfterMs", () => {
  it("parses delay-seconds", () => {
    expect(parseRetryAfterMs("5")).toBe(5_000)
    expect(parseRetryAfterMs("  12 ")).toBe(12_000)
    expect(parseRetryAfterMs("0")).toBe(0)
  })

  it("caps at 60 s", () => {
    expect(parseRetryAfterMs("600")).toBe(60_000)
  })

  it("returns null for absent or unparseable values", () => {
    expect(parseRetryAfterMs(null)).toBeNull()
    expect(parseRetryAfterMs(undefined)).toBeNull()
    expect(parseRetryAfterMs("")).toBeNull()
    expect(parseRetryAfterMs("   ")).toBeNull()
    expect(parseRetryAfterMs("soon")).toBeNull()
  })

  // "5.5" would otherwise reach Date.parse and yield an ambiguous result.
  it("rejects bare non-integer numerics rather than date-parsing them", () => {
    expect(parseRetryAfterMs("5.5")).toBeNull()
  })

  it("parses an HTTP-date into a delta", () => {
    const future = new Date(Date.now() + 10_000).toUTCString()
    const parsed = parseRetryAfterMs(future)
    expect(parsed).not.toBeNull()
    expect(parsed!).toBeGreaterThan(5_000)
    expect(parsed!).toBeLessThanOrEqual(10_000)
  })

  it("clamps a past HTTP-date to zero", () => {
    expect(parseRetryAfterMs(new Date(Date.now() - 60_000).toUTCString())).toBe(0)
  })
})

describe("backoffDelayMs", () => {
  it("grows exponentially and respects the ceiling", () => {
    const full = () => 0.999999
    expect(backoffDelayMs(0, 100, 5_000, full)).toBe(99)
    expect(backoffDelayMs(1, 100, 5_000, full)).toBe(199)
    expect(backoffDelayMs(2, 100, 5_000, full)).toBe(399)
    expect(backoffDelayMs(20, 100, 5_000, full)).toBe(4_999)
  })

  it("applies full jitter — a zero roll yields no delay", () => {
    expect(backoffDelayMs(5, 100, 5_000, () => 0)).toBe(0)
  })
})

describe("isMethodIdempotent", () => {
  it("classifies per RFC 7231", () => {
    for (const m of ["GET", "HEAD", "PUT", "DELETE", "OPTIONS", "TRACE", "get", "delete"]) {
      expect(isMethodIdempotent(m)).toBe(true)
    }
    for (const m of ["POST", "PATCH", "post", "patch"]) {
      expect(isMethodIdempotent(m)).toBe(false)
    }
  })
})

describe("CircuitBreaker", () => {
  it("opens after the threshold and reports remaining cooldown", () => {
    let now = 1_000
    const breaker = new CircuitBreaker(3, 30_000, false, () => now)

    expect(breaker.checkBeforeRequest()).toBeNull()
    breaker.recordFailure()
    breaker.recordFailure()
    expect(breaker.checkBeforeRequest()).toBeNull()

    breaker.recordFailure()
    expect(breaker.getState().open).toBe(true)
    expect(breaker.checkBeforeRequest()).toBe(30_000)

    now += 10_000
    expect(breaker.checkBeforeRequest()).toBe(20_000)
  })

  it("half-opens once the cooldown elapses", () => {
    let now = 0
    const breaker = new CircuitBreaker(1, 5_000, false, () => now)
    breaker.recordFailure()
    expect(breaker.checkBeforeRequest()).toBe(5_000)

    now = 5_000
    expect(breaker.checkBeforeRequest()).toBeNull()
  })

  it("a failed probe extends the cooldown", () => {
    let now = 0
    const breaker = new CircuitBreaker(1, 5_000, false, () => now)
    breaker.recordFailure()

    now = 5_000
    expect(breaker.checkBeforeRequest()).toBeNull()
    breaker.recordFailure()
    expect(breaker.checkBeforeRequest()).toBe(5_000)
  })

  it("a successful probe closes the breaker", () => {
    let now = 0
    const breaker = new CircuitBreaker(1, 5_000, false, () => now)
    breaker.recordFailure()

    now = 5_000
    breaker.recordSuccess()
    expect(breaker.getState()).toMatchObject({ failures: 0, openedAt: null, open: false })
  })

  it("stays closed when disabled", () => {
    const breaker = new CircuitBreaker(1, 5_000, true)
    breaker.recordFailure()
    breaker.recordFailure()
    expect(breaker.checkBeforeRequest()).toBeNull()
  })
})

describe("fetchWithResilience", () => {
  beforeEach(() => vi.restoreAllMocks())

  it("returns a successful response without retrying", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(res(200)))
    const response = await fetchWithResilience("https://g", undefined, config(), openBreaker(), noSleep, fetchImpl)

    expect(response.status).toBe(200)
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it("does not retry a deterministic 4xx", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(res(403)))
    const response = await fetchWithResilience("https://g", undefined, config(), openBreaker(), noSleep, fetchImpl)

    expect(response.status).toBe(403)
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it("retries a 429 and returns the eventual success", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(res(429))
      .mockResolvedValueOnce(res(429))
      .mockResolvedValueOnce(res(200))

    const response = await fetchWithResilience("https://g", undefined, config(), openBreaker(), noSleep, fetchImpl)
    expect(response.status).toBe(200)
    expect(fetchImpl).toHaveBeenCalledTimes(3)
  })

  it("honours Retry-After on a 429", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(res(429, { "retry-after": "2" }))
      .mockResolvedValueOnce(res(200))
    const sleep = vi.fn(() => Promise.resolve())

    await fetchWithResilience("https://g", undefined, config(), openBreaker(), sleep, fetchImpl)
    expect(sleep).toHaveBeenCalledWith(2_000)
  })

  // Graph throttles before executing, so the side effect has not landed.
  it("retries a throttled POST", async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(res(429)).mockResolvedValueOnce(res(201))
    const response = await fetchWithResilience(
      "https://g",
      { method: "POST" },
      config(),
      openBreaker(),
      noSleep,
      fetchImpl,
    )

    expect(response.status).toBe(201)
    expect(fetchImpl).toHaveBeenCalledTimes(2)
  })

  it("retries a 503 for an idempotent method", async () => {
    const fetchImpl = vi.fn().mockResolvedValueOnce(res(503)).mockResolvedValueOnce(res(200))
    const response = await fetchWithResilience(
      "https://g",
      { method: "GET" },
      config(),
      openBreaker(),
      noSleep,
      fetchImpl,
    )

    expect(response.status).toBe(200)
    expect(fetchImpl).toHaveBeenCalledTimes(2)
  })

  // The request may already have executed server-side; a retry would duplicate it.
  it("does NOT retry a 503 for POST", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(res(503)))
    const response = await fetchWithResilience(
      "https://g",
      { method: "POST" },
      config(),
      openBreaker(),
      noSleep,
      fetchImpl,
    )

    expect(response.status).toBe(503)
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it("does NOT retry a network error for PATCH", async () => {
    const fetchImpl = vi.fn(() => Promise.reject(new Error("socket hang up")))
    await expect(
      fetchWithResilience("https://g", { method: "PATCH" }, config(), openBreaker(), noSleep, fetchImpl),
    ).rejects.toThrow("socket hang up")
    expect(fetchImpl).toHaveBeenCalledTimes(1)
  })

  it("retries a network error for GET and rethrows once the budget is spent", async () => {
    const fetchImpl = vi.fn(() => Promise.reject(new Error("ECONNRESET")))
    await expect(
      fetchWithResilience("https://g", { method: "GET" }, config({ maxRetries: 2 }), openBreaker(), noSleep, fetchImpl),
    ).rejects.toThrow("ECONNRESET")

    expect(fetchImpl).toHaveBeenCalledTimes(3) // initial + 2 retries
  })

  it("surfaces the final response when retries are exhausted", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(res(429)))
    const response = await fetchWithResilience(
      "https://g",
      undefined,
      config({ maxRetries: 2 }),
      openBreaker(),
      noSleep,
      fetchImpl,
    )

    expect(response.status).toBe(429)
    expect(fetchImpl).toHaveBeenCalledTimes(3)
  })

  it("drains a retried response body so the socket is not leaked", async () => {
    const drained = res(429)
    const arrayBuffer = vi.spyOn(drained, "arrayBuffer")
    const fetchImpl = vi.fn().mockResolvedValueOnce(drained).mockResolvedValueOnce(res(200))

    await fetchWithResilience("https://g", undefined, config(), openBreaker(), noSleep, fetchImpl)
    expect(arrayBuffer).toHaveBeenCalled()
  })

  it("passes an abort signal so a stuck call cannot hang forever", async () => {
    const fetchImpl = vi.fn(() => Promise.resolve(res(200)))
    await fetchWithResilience("https://g", { method: "GET" }, config(), openBreaker(), noSleep, fetchImpl)

    const init = fetchImpl.mock.calls[0]![1] as RequestInit
    expect(init.signal).toBeInstanceOf(AbortSignal)
  })

  it("fast-fails while the circuit is open", async () => {
    const breaker = new CircuitBreaker(1, 30_000, false)
    breaker.recordFailure()
    const fetchImpl = vi.fn(() => Promise.resolve(res(200)))

    await expect(
      fetchWithResilience("https://g", undefined, config(), breaker, noSleep, fetchImpl),
    ).rejects.toBeInstanceOf(CircuitOpenError)
    expect(fetchImpl).not.toHaveBeenCalled()
  })

  it("counts repeated failures toward opening the circuit", async () => {
    const breaker = new CircuitBreaker(2, 30_000, false)
    const fetchImpl = vi.fn(() => Promise.resolve(res(503)))

    // Each exhausted GET records one failure.
    await fetchWithResilience("https://g", { method: "GET" }, config({ maxRetries: 0 }), breaker, noSleep, fetchImpl)
    expect(breaker.getState().open).toBe(false)
    await fetchWithResilience("https://g", { method: "GET" }, config({ maxRetries: 0 }), breaker, noSleep, fetchImpl)

    expect(breaker.getState().open).toBe(true)
  })
})

describe("isCircuitOpenError", () => {
  it("narrows only its own error type", () => {
    expect(isCircuitOpenError(new CircuitOpenError(5_000))).toBe(true)
    expect(isCircuitOpenError(new Error("nope"))).toBe(false)
    expect(isCircuitOpenError(null)).toBe(false)
  })
})

describe("loadResilienceConfig", () => {
  it("falls back to documented defaults", () => {
    const cfg = loadResilienceConfig()
    expect(cfg).toMatchObject({ maxRetries: 3, fetchTimeoutMs: 100_000, circuitFailureThreshold: 5 })
  })

  it("reads overrides from the environment", () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "7")
    vi.stubEnv("MS365_GRAPH_CIRCUIT_DISABLED", "true")
    const cfg = loadResilienceConfig()

    expect(cfg.maxRetries).toBe(7)
    expect(cfg.circuitDisabled).toBe(true)
    vi.unstubAllEnvs()
  })

  it("ignores a non-numeric override rather than propagating NaN", () => {
    vi.stubEnv("MS365_GRAPH_MAX_RETRIES", "lots")
    vi.spyOn(console, "error").mockImplementation(() => {})

    expect(loadResilienceConfig().maxRetries).toBe(3)
    vi.unstubAllEnvs()
  })
})
