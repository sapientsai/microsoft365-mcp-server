// The decorator around token acquisition: it detects a token that predates a permission
// change and says so, without touching the token request itself.
//
// The "never sends a claims challenge" test is a regression guard, not a style rule. An
// earlier version sent one to force a fresh token; Azure validates the claims payload and
// rejected it with AADSTS1000004, which fails the whole authorization request and leaves
// the user unable to sign in at all.

import type { AccessToken, TokenCredential } from "@azure/identity"
import { beforeEach, describe, expect, it, vi } from "vitest"

import { revalidateScopes } from "../src/auth/auth-modes"

const REQUIRED = ["Mail.ReadWrite.Shared"]

const accessToken = (token: string): AccessToken => ({ token, expiresOnTimestamp: Date.now() + 3_600_000 })

// Stands in for the JWT reader: the token string is its own scope list, so these tests
// exercise the decorator's control flow rather than jwt.decode.
const readScopes = (token: string): ReadonlyArray<string> => (token.length === 0 ? [] : token.split(" "))

let consoleError: ReturnType<typeof vi.spyOn>

beforeEach(() => {
  // Restored between tests: a spy that survives would carry the previous test's calls
  // into the "reported only once" count and make it pass or fail for the wrong reason.
  vi.restoreAllMocks()
  consoleError = vi.spyOn(console, "error").mockImplementation(() => {})
})

describe("revalidateScopes", () => {
  it("passes the credential through untouched when nothing is required", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, [], readScopes)

    expect(credential).toBe(inner)
  })

  it("returns the token unchanged when every required scope is present", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite Mail.ReadWrite.Shared")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    const token = await credential.getToken("scope")

    expect(token?.token).toBe("Mail.ReadWrite Mail.ReadWrite.Shared")
    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })

  it("reports the shortfall when the token predates a permission change", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    expect(consoleError).toHaveBeenCalledWith(expect.stringContaining("Mail.ReadWrite.Shared"))
  })

  // AADSTS1000004: Azure validates the claims payload and rejects a value it does not
  // recognise, failing the entire authorization request. Detection must never alter the
  // token request — a diagnostic that can break sign-in is worse than the problem.
  it("never sends a claims challenge", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    for (const [, options] of inner.getToken.mock.calls as Array<[unknown, { claims?: string } | undefined]>) {
      expect(options?.claims).toBeUndefined()
    }
  })

  it("acquires the token exactly once, drift or not", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })

  it("passes the caller's options through untouched", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)
    const options = { tenantId: "t" }

    await credential.getToken("scope", options)

    expect(inner.getToken).toHaveBeenCalledWith("scope", options)
  })

  it("names both the permission grant and the stale cache in the message", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    const [line] = consoleError.mock.calls[0] as [string]
    expect(line).toContain("app registration")
    expect(line).toContain("token cache")
  })

  it("still returns a usable token when the scope is missing, rather than failing the call", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    const token = await credential.getToken("scope")

    // Graph refuses the calls that need the scope; everything else keeps working.
    expect(token?.token).toBe("Mail.ReadWrite")
  })

  // This runs on every token acquisition, so an unfixable shortfall must not print a line
  // per Graph call.
  it("reports an unfixable shortfall only once", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("Mail.ReadWrite")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")
    await credential.getToken("scope")
    await credential.getToken("scope")

    const drift = consoleError.mock.calls.filter(([line]) => String(line).includes("app registration"))
    expect(drift).toHaveLength(1)
  })

  it("leaves a null token alone", async () => {
    const inner = { getToken: vi.fn(async () => null) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    expect(await credential.getToken("scope")).toBeNull()
    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })

  // An opaque token reads as no scopes at all; treating that as drift would re-acquire on
  // every single call for deployments this cannot parse.
  it("does not re-acquire when the token's scopes cannot be read", async () => {
    const inner = { getToken: vi.fn(async () => accessToken("")) }
    const credential = revalidateScopes(inner as unknown as TokenCredential, REQUIRED, readScopes)

    await credential.getToken("scope")

    expect(inner.getToken).toHaveBeenCalledTimes(1)
  })
})
