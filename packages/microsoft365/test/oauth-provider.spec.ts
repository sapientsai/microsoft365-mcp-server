import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { resolveEncryptionKey, resolveSigningKey, resolveTokenStoragePath } from "../src/auth/oauth-provider"
import { DEFAULT_INTERACTIVE_SCOPES, resolveInteractiveScopes } from "../src/auth/scopes"

const KEYS = [
  "MS365_JWT_SIGNING_KEY",
  "MS365_TOKEN_ENCRYPTION_KEY",
  "TOKEN_STORAGE_PATH",
  "MS365_EXTRA_SCOPES",
] as const

describe("oauth-provider key separation", () => {
  beforeEach(() => {
    vi.restoreAllMocks()
    vi.spyOn(console, "error").mockImplementation(() => {})
    for (const k of KEYS) delete process.env[k]
  })
  afterEach(() => {
    for (const k of KEYS) delete process.env[k]
  })

  describe("resolveSigningKey", () => {
    it("prefers the dedicated MS365_JWT_SIGNING_KEY", () => {
      process.env.MS365_JWT_SIGNING_KEY = "dedicated-signing"
      expect(resolveSigningKey("client-secret")).toBe("dedicated-signing")
      expect(console.error).not.toHaveBeenCalled()
    })
    it("falls back to the client secret and warns when unset", () => {
      expect(resolveSigningKey("client-secret")).toBe("client-secret")
      expect(console.error).toHaveBeenCalledWith(expect.stringContaining("MS365_JWT_SIGNING_KEY"))
    })
  })

  describe("resolveEncryptionKey", () => {
    it("prefers the dedicated MS365_TOKEN_ENCRYPTION_KEY", () => {
      process.env.MS365_TOKEN_ENCRYPTION_KEY = "dedicated-enc"
      expect(resolveEncryptionKey("client-secret")).toBe("dedicated-enc")
      expect(console.error).not.toHaveBeenCalled()
    })
    it("falls back to the client secret and warns when unset", () => {
      expect(resolveEncryptionKey("client-secret")).toBe("client-secret")
      expect(console.error).toHaveBeenCalledWith(expect.stringContaining("MS365_TOKEN_ENCRYPTION_KEY"))
    })
  })

  describe("resolveTokenStoragePath", () => {
    it("honors TOKEN_STORAGE_PATH", () => {
      process.env.TOKEN_STORAGE_PATH = "/data/tokens"
      expect(resolveTokenStoragePath()).toBe("/data/tokens")
    })
    // This default is load-bearing in production, which is not obvious from reading it.
    //
    // The deployed ms365 service does not set TOKEN_STORAGE_PATH. It persists tokens by mounting
    // a named volume at /tmp/ms365-tokens — this exact string. The two are coupled by
    // coincidence, and changing this line quietly breaks that mount: every user is logged out
    // with "requires re-authorization" and no error appears anywhere, because nothing failed.
    //
    // So this is not a test of a default, it is a tripwire on a deployment contract. If the
    // default genuinely needs to move, the volume mount has to move with it — or better, set
    // TOKEN_STORAGE_PATH explicitly in the deployment first and make the coupling real.
    it("defaults to the path production mounts a volume at", () => {
      expect(
        resolveTokenStoragePath(),
        "A deployed service mounts its token volume at this exact path without setting " +
          "TOKEN_STORAGE_PATH. Changing this default logs every user out silently. Move the " +
          "deployment to an explicit TOKEN_STORAGE_PATH before changing it here.",
      ).toBe("/tmp/ms365-tokens")
    })
  })

  // The scope set the provider requests is what ends up in the issued access token. Before this,
  // createAzureAuthProvider's only call site never passed `scopes`, so DEFAULT_INTERACTIVE_SCOPES
  // was effectively hardcoded and no admin grant could reach the token without a release.
  describe("MS365_EXTRA_SCOPES", () => {
    it("is read from the environment, not just from an argument", () => {
      process.env.MS365_EXTRA_SCOPES = "OnlineMeetingTranscript.Read.All"

      expect(resolveInteractiveScopes()).toContain("OnlineMeetingTranscript.Read.All")
    })

    it("leaves the requested set untouched when unset", () => {
      expect(resolveInteractiveScopes()).toEqual([...DEFAULT_INTERACTIVE_SCOPES])
    })
  })
})
