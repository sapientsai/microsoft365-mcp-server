import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { resolveEncryptionKey, resolveSigningKey, resolveTokenStoragePath } from "../src/auth/oauth-provider"

const KEYS = ["MS365_JWT_SIGNING_KEY", "MS365_TOKEN_ENCRYPTION_KEY", "TOKEN_STORAGE_PATH"] as const

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
    it("defaults when unset", () => {
      expect(resolveTokenStoragePath()).toBe("/tmp/ms365-tokens")
    })
  })
})
