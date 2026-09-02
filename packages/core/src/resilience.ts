/**
 * Resilience layer for Microsoft Graph calls.
 *
 * Ported from Softeria/ms-365-mcp-server (`src/lib/graph-resilience.ts`, MIT),
 * adapted to this workspace: `MS365_`-prefixed env vars, `console.error` logging
 * (stdio transport keeps stdout clear for the protocol), and a `Left`-returning
 * circuit-open path rather than a thrown error.
 *
 * Three concerns folded into one module:
 *
 *  1. **Fetch timeout** via AbortController — a stuck Graph call must not hang an
 *     MCP request indefinitely. Default 100 s (matches the .NET `HttpClient` /
 *     `aiohttp` defaults so large direct uploads of 50–250 MB over slow links
 *     don't get aborted mid-flight). Override with `MS365_GRAPH_TIMEOUT_MS`.
 *
 *  2. **Retry with backoff** on transient failures:
 *       - HTTP 429 — honour `Retry-After` (seconds or HTTP-date), cap at 60 s.
 *         Safe to retry on every method including POST / PATCH because Graph
 *         throttles *before* executing the operation; the side effect has not
 *         landed server-side when a 429 comes back.
 *       - HTTP 503 / 504 / network errors — retried for **idempotent methods
 *         only** (GET / HEAD / PUT / DELETE / OPTIONS / TRACE, per RFC 7231).
 *         POST and PATCH cannot be retried on these failures: a client-side
 *         timeout or 5xx after the request was already executing server-side
 *         would silently duplicate the side effect.
 *       - Everything else (auth, invalid input, 403 scope errors) — NOT retried.
 *         Those are deterministic.
 *     Default 3 retries, override with `MS365_GRAPH_MAX_RETRIES`.
 *
 *  3. **Circuit breaker** — a process-wide singleton tracks consecutive failures
 *     against Graph. After `MS365_GRAPH_CIRCUIT_THRESHOLD` failures (default 5)
 *     the breaker opens and every subsequent call fast-fails for
 *     `MS365_GRAPH_CIRCUIT_COOLDOWN_MS` (default 30 s) before half-opening for a
 *     probe. Prevents flooding Graph when it is already on fire. Disable with
 *     `MS365_GRAPH_CIRCUIT_DISABLED=true`.
 *
 * All knobs are env-var-driven so a deployment can be tuned without a code change.
 */

export type ResilienceConfig = {
  readonly maxRetries: number
  readonly baseBackoffMs: number
  readonly maxBackoffMs: number
  readonly fetchTimeoutMs: number
  readonly circuitFailureThreshold: number
  readonly circuitCooldownMs: number
  readonly circuitDisabled: boolean
}

const intEnv = (name: string, fallback: number): number => {
  const raw = process.env[name]
  if (raw === undefined || raw === "") return fallback
  const n = Number.parseInt(raw, 10)
  if (!Number.isFinite(n) || n < 0) {
    console.error(`[Resilience] Ignoring invalid ${name}=${JSON.stringify(raw)} (use a non-negative integer)`)
    return fallback
  }
  return n
}

export const loadResilienceConfig = (): ResilienceConfig => ({
  maxRetries: intEnv("MS365_GRAPH_MAX_RETRIES", 3),
  baseBackoffMs: intEnv("MS365_GRAPH_BASE_BACKOFF_MS", 200),
  maxBackoffMs: intEnv("MS365_GRAPH_MAX_BACKOFF_MS", 5_000),
  fetchTimeoutMs: intEnv("MS365_GRAPH_TIMEOUT_MS", 100_000),
  circuitFailureThreshold: intEnv("MS365_GRAPH_CIRCUIT_THRESHOLD", 5),
  circuitCooldownMs: intEnv("MS365_GRAPH_CIRCUIT_COOLDOWN_MS", 30_000),
  circuitDisabled:
    process.env.MS365_GRAPH_CIRCUIT_DISABLED === "true" || process.env.MS365_GRAPH_CIRCUIT_DISABLED === "1",
})

export class CircuitBreaker {
  private failures = 0
  private openedAt: number | null = null

  constructor(
    private readonly threshold: number,
    private readonly cooldownMs: number,
    private readonly disabled: boolean,
    private readonly now: () => number = () => Date.now(),
  ) {}

  /**
   * @returns ms remaining before the circuit can be probed, or `null` if the
   *          circuit is closed and the call should proceed.
   */
  checkBeforeRequest(): number | null {
    if (this.disabled) return null
    if (this.openedAt === null) return null
    const elapsed = this.now() - this.openedAt
    // Half-open — let one probe through; success closes the circuit, failure
    // resets the cooldown timer.
    if (elapsed >= this.cooldownMs) return null
    return this.cooldownMs - elapsed
  }

  recordSuccess(): void {
    if (this.failures !== 0 || this.openedAt !== null) {
      console.error("[Resilience] Graph circuit: success — closing breaker")
    }
    this.failures = 0
    this.openedAt = null
  }

  recordFailure(): void {
    if (this.disabled) return
    this.failures += 1
    if (this.failures >= this.threshold && this.openedAt === null) {
      this.openedAt = this.now()
      console.error(
        `[Resilience] Graph circuit: ${this.failures} consecutive failures — opening breaker for ${this.cooldownMs} ms`,
      )
    } else if (this.openedAt !== null) {
      // Failed during the probe → reset the cooldown clock.
      this.openedAt = this.now()
      console.error("[Resilience] Graph circuit: probe failed — extending cooldown")
    }
  }

  /** Exposed for tests / metrics. */
  getState(): { failures: number; openedAt: number | null; open: boolean } {
    return { failures: this.failures, openedAt: this.openedAt, open: this.checkBeforeRequest() !== null }
  }
}

/**
 * Parse a Retry-After header (seconds or HTTP-date). Returns null if absent or
 * unparseable. Caps the delay at 60 s — beyond that we'd rather surface the
 * throttle to the caller than hang the connection.
 */
export const parseRetryAfterMs = (header: string | null | undefined): number | null => {
  if (!header) return null
  const trimmed = header.trim()
  if (trimmed === "") return null

  const asInt = Number.parseInt(trimmed, 10)
  if (Number.isFinite(asInt) && asInt >= 0 && String(asInt) === trimmed) {
    return Math.min(asInt * 1000, 60_000)
  }

  // HTTP-date branch — require a credible date-shaped string. RFC 7231
  // IMF-fixdate / obs-date / ANSI C all contain at least one of these delimiters,
  // while bare numerics like "5.5" would otherwise be parsed as ambiguous Date
  // inputs on some Node versions.
  if (!/[-/:,]| GMT$/i.test(trimmed) && !/\s+\d/.test(trimmed)) return null

  const dateMs = Date.parse(trimmed)
  if (Number.isFinite(dateMs)) {
    const delta = dateMs - Date.now()
    if (delta <= 0) return 0
    return Math.min(delta, 60_000)
  }
  return null
}

// Exponential backoff with full jitter: random in [0, min(max, base * 2^attempt))
export const backoffDelayMs = (
  attempt: number,
  baseMs: number,
  maxMs: number,
  rand: () => number = Math.random,
): number => Math.floor(rand() * Math.min(maxMs, baseMs * 2 ** attempt))

const isRetriableStatus = (status: number): boolean => status === 429 || status === 503 || status === 504

/**
 * Per RFC 7231 §4.2.2 (and RFC 5789 §1 for PATCH): GET, HEAD, PUT, DELETE,
 * OPTIONS, TRACE are idempotent — retrying them after a network failure or 5xx is
 * safe because applying the request N times has the same effect as applying it
 * once. POST and PATCH are explicitly NOT idempotent. 429 is still safe to retry
 * on any method because the throttling decision happens before Graph executes the
 * operation.
 */
export const isMethodIdempotent = (method: string): boolean => {
  const m = method.toUpperCase()
  return m === "GET" || m === "HEAD" || m === "PUT" || m === "DELETE" || m === "OPTIONS" || m === "TRACE"
}

const isAbortError = (err: unknown): boolean =>
  typeof err === "object" && err !== null && "name" in err && (err as { name: string }).name === "AbortError"

export class CircuitOpenError extends Error {
  readonly code = "circuit_open"
  readonly cooldownMs: number
  constructor(cooldownMs: number) {
    super(
      `Graph circuit breaker is open (cooldown ${cooldownMs} ms). Upstream has failed repeatedly; refusing to flood it.`,
    )
    this.name = "CircuitOpenError"
    this.cooldownMs = cooldownMs
  }
}

export const isCircuitOpenError = (err: unknown): err is CircuitOpenError => err instanceof CircuitOpenError

export type FetchLike = (url: string, init?: RequestInit) => Promise<Response>

/**
 * Wraps `fetch` with timeout + retry + circuit-breaker semantics.
 *
 * The signature mirrors `fetch` so it drops into existing call sites: pass the URL
 * and `init`, get back a `Response`. On retriable failure exhausting the budget,
 * the final attempt's Response (or thrown error) is surfaced unchanged — callers
 * handle 429 / 5xx exactly as they did before.
 *
 * Throws `CircuitOpenError` when the breaker is open.
 */
export const fetchWithResilience = async (
  url: string,
  init: RequestInit | undefined,
  config: ResilienceConfig,
  breaker: CircuitBreaker,
  sleep: (ms: number) => Promise<void> = (ms) => new Promise((r) => setTimeout(r, ms)),
  fetchImpl: FetchLike = fetch,
): Promise<Response> => {
  const remainingCooldown = breaker.checkBeforeRequest()
  if (remainingCooldown !== null) throw new CircuitOpenError(remainingCooldown)

  const method = (init?.method ?? "GET").toString().toUpperCase()
  const methodIsIdempotent = isMethodIdempotent(method)

  /* eslint-disable functype/no-let -- retry loop: the attempt counter advances per
     iteration, and the fetch outcome has to escape the try/catch that produced it. */
  let attempt = 0
  for (;;) {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), config.fetchTimeoutMs)

    let response: Response | null = null
    let networkError: unknown = null
    try {
      response = await fetchImpl(url, { ...init, signal: controller.signal })
    } catch (err) {
      networkError = err
    } finally {
      clearTimeout(timer)
    }

    if (response !== null && !isRetriableStatus(response.status)) {
      breaker.recordSuccess()
      return response
    }

    // Non-idempotent methods (POST, PATCH) cannot be safely retried on
    // 503/504/network errors — the request may have already executed server-side.
    // 429 is the exception: Graph throttles before executing, so retrying a
    // throttled POST is safe and follows Graph's documented contract.
    const is429 = response !== null && response.status === 429
    const retryAllowedByMethod = methodIsIdempotent || is429
    const canRetry = attempt < config.maxRetries && retryAllowedByMethod

    if (!canRetry) {
      breaker.recordFailure()
      if (!retryAllowedByMethod && attempt === 0) {
        const what = response !== null ? `${response.status}` : "network error"
        console.error(
          `[Resilience] Graph ${method} ${what}: not retried (non-idempotent method, side-effect may have landed)`,
        )
      }
      if (response !== null) return response
      throw networkError ?? new Error("Graph fetch failed (unknown error)")
    }

    const delayMs =
      response !== null && response.status === 429
        ? (parseRetryAfterMs(response.headers.get("retry-after")) ??
          backoffDelayMs(attempt, config.baseBackoffMs, config.maxBackoffMs))
        : backoffDelayMs(attempt, config.baseBackoffMs, config.maxBackoffMs)

    const reason =
      response !== null
        ? `HTTP ${response.status}`
        : isAbortError(networkError)
          ? `timeout (${config.fetchTimeoutMs} ms)`
          : `network error: ${networkError instanceof Error ? networkError.message : "unknown"}`
    console.error(
      `[Resilience] Graph retry ${attempt + 1}/${config.maxRetries} after ${reason} — sleeping ${delayMs} ms`,
    )

    // Drain the body so we don't leak the underlying socket.
    if (response !== null) {
      try {
        await response.arrayBuffer()
      } catch {
        // Best-effort cleanup.
      }
    }

    breaker.recordFailure()
    attempt += 1
    await sleep(delayMs)
  }
  /* eslint-enable functype/no-let */
}

// Process-wide breaker. Tests pass their own breaker into fetchWithResilience.
// eslint-disable-next-line functype/no-let -- lazily built singleton, reset by tests
let sharedBreaker: CircuitBreaker | null = null

export const getSharedBreaker = (): CircuitBreaker => {
  if (sharedBreaker === null) {
    const cfg = loadResilienceConfig()
    sharedBreaker = new CircuitBreaker(cfg.circuitFailureThreshold, cfg.circuitCooldownMs, cfg.circuitDisabled)
  }
  return sharedBreaker
}

export const resetSharedBreakerForTests = (): void => {
  sharedBreaker = null
}

/**
 * A `fetch`-shaped function with resilience applied, using the process-wide
 * breaker and env-driven config. Config is read once per call so an env change in
 * tests takes effect without a module reload.
 */
export const resilientFetch: FetchLike = (url, init) =>
  fetchWithResilience(url, init, loadResilienceConfig(), getSharedBreaker())
