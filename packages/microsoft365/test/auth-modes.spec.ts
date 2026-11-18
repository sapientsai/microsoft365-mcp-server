import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

import {
  type AuthenticationRecord,
  AuthenticationRequiredError,
  DeviceCodeCredential,
  InteractiveBrowserCredential,
  type TokenCredential,
} from "@azure/identity"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { createCredential } from "../src/auth/auth-modes"
import { createFileCachePlugin, readAuthenticationRecord, writeAuthenticationRecord } from "../src/auth/token-cache"

const config = { mode: "interactive" as const, tenantId: "common", clientId: "11111111-1111-1111-1111-111111111111" }
const token = { token: "test-token", expiresOnTimestamp: 1_000 }

// Keep the persisted token cache and authentication record out of the real config directory.
const cacheDir = { path: "" }

beforeEach(() => {
  cacheDir.path = mkdtempSync(join(tmpdir(), "ms365-auth-modes-"))
  vi.stubEnv("MS365_TOKEN_CACHE_PATH", cacheDir.path)
})

afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllEnvs()
  rmSync(cacheDir.path, { recursive: true, force: true })
})

describe("interactive authentication mode", () => {
  it("uses the browser by default", async () => {
    const browserToken = vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockResolvedValue(token)
    const result = createCredential(config)
    expect(result.isRight()).toBe(true)

    // With the token cache on (the default) the browser credential is wrapped so the
    // authentication record can be saved after sign-in; it still does the signing in.
    await expect((result.value as TokenCredential).getToken("scope")).resolves.toBe(token)
    expect(browserToken).toHaveBeenCalledTimes(1)
  })

  it("hands back the bare browser credential when the token cache is off", () => {
    vi.stubEnv("MS365_TOKEN_CACHE", "false")
    const result = createCredential(config)
    expect(result.isRight()).toBe(true)
    expect(result.value).toBeInstanceOf(InteractiveBrowserCredential)
  })

  it("uses one device-code credential for repeated token requests when explicitly enabled", async () => {
    vi.stubEnv("MS365_TOKEN_CACHE", "false")
    const deviceToken = vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
    const browserToken = vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
    const result = createCredential({ ...config, useDeviceCode: true })
    expect(result.isRight()).toBe(true)
    const credential = result.value as TokenCredential
    expect(credential).toBeInstanceOf(DeviceCodeCredential)

    const scopes = ["https://graph.microsoft.com/.default"]
    await expect(credential.getToken(scopes)).resolves.toBe(token)
    await expect(credential.getToken(scopes)).resolves.toBe(token)
    expect(deviceToken).toHaveBeenCalledTimes(2)
    expect(deviceToken.mock.contexts).toEqual([credential, credential])
    expect(browserToken).not.toHaveBeenCalled()
  })

  describe("after a configuration change", () => {
    const seedCache = async () =>
      createFileCachePlugin("microsoft365-mcp-server", cacheDir.path).afterCacheAccess({
        cacheHasChanged: true,
        tokenCache: {
          serialize: () => JSON.stringify({ AccessToken: { at: {} }, RefreshToken: { rt: {} } }),
          deserialize: () => undefined,
        },
      } as never)
    const cached = () => {
      const file = readdirSync(cacheDir.path).find((f) => f.startsWith("microsoft365-mcp-server."))!
      return JSON.parse(readFileSync(join(cacheDir.path, file), "utf8")) as Record<string, object>
    }

    it("drops the cached access token and keeps the refresh token", async () => {
      vi.spyOn(console, "error").mockImplementation(() => undefined)
      vi.stubEnv("MS365_ALLOWED_MAILBOXES", "a@example.com")
      createCredential(config)
      await seedCache()

      vi.stubEnv("MS365_ALLOWED_MAILBOXES", "a@example.com,b@example.com")
      createCredential(config)
      expect(cached()).toEqual({ AccessToken: {}, RefreshToken: { rt: {} } })
    })

    it("keeps the cached access token when nothing changed", async () => {
      createCredential(config)
      await seedCache()
      createCredential(config)
      expect(cached().AccessToken).toEqual({ at: {} })
    })

    it("does not touch the disk when the token cache is off", () => {
      vi.stubEnv("MS365_TOKEN_CACHE", "false")
      createCredential(config)
      expect(existsSync(join(cacheDir.path, "config-fingerprint"))).toBe(false)
    })
  })

  it("propagates browser errors without starting another authentication flow", async () => {
    const error = new Error("Unable to open browser")
    vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockRejectedValue(error)
    const deviceToken = vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
    const result = createCredential(config)
    const credential = result.value as TokenCredential

    await expect(credential.getToken("scope")).rejects.toBe(error)
    expect(deviceToken).not.toHaveBeenCalled()
  })

  describe("signing in when the cache cannot serve a token", () => {
    const record: AuthenticationRecord = {
      authority: "https://login.microsoftonline.com/common",
      homeAccountId: "uid.utid",
      tenantId: "utid",
      username: "someone@example.com",
      clientId: config.clientId,
    }
    const required = () => new AuthenticationRequiredError({ scopes: ["scope"] })

    it("signs in through authenticate() and persists the record it returns", async () => {
      const getToken = vi
        .spyOn(InteractiveBrowserCredential.prototype, "getToken")
        .mockRejectedValueOnce(required())
        .mockResolvedValue(token)
      const authenticate = vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate").mockResolvedValue(record)
      const credential = createCredential(config).value as TokenCredential

      await expect(credential.getToken("scope")).resolves.toBe(token)

      expect(authenticate).toHaveBeenCalledOnce()
      expect(getToken).toHaveBeenCalledTimes(2)
      expect(readAuthenticationRecord(cacheDir.path)).toEqual(record)
    })

    it("does not open the browser when the cache serves the token", async () => {
      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockResolvedValue(token)
      const authenticate = vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate")
      const credential = createCredential(config).value as TokenCredential

      await expect(credential.getToken("scope")).resolves.toBe(token)
      expect(authenticate).not.toHaveBeenCalled()
    })

    // Parallel Graph calls can all miss the cache at once; each opening a browser window would be absurd.
    it("shares one sign-in between concurrent requests", async () => {
      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
        .mockRejectedValueOnce(required())
        .mockRejectedValueOnce(required())
        .mockResolvedValue(token)
      const authenticate = vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate").mockResolvedValue(record)
      const credential = createCredential(config).value as TokenCredential

      await Promise.all([credential.getToken("scope"), credential.getToken("scope")])
      expect(authenticate).toHaveBeenCalledOnce()
    })

    it("says so when a saved sign-in could not be reused, and stays quiet on a first sign-in", async () => {
      const log = vi.spyOn(console, "error").mockImplementation(() => undefined)
      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
        .mockRejectedValueOnce(required())
        .mockResolvedValue(token)
      vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate").mockResolvedValue(record)

      await (createCredential(config).value as TokenCredential).getToken("scope")
      expect(log).not.toHaveBeenCalled()

      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
        .mockRejectedValueOnce(required())
        .mockResolvedValue(token)
      await (createCredential(config).value as TokenCredential).getToken("scope")
      expect(log).toHaveBeenCalledWith(expect.stringContaining("Cached credentials could not be reused"))
    })

    it("explains a failure after sign-in instead of passing on identity's advice to call authenticate()", async () => {
      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
        .mockRejectedValueOnce(required())
        .mockRejectedValueOnce(required())
      vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate").mockResolvedValue(record)
      const credential = createCredential(config).value as TokenCredential

      const failure = await credential.getToken("scope").catch((e: unknown) => e)
      expect(failure).toBeInstanceOf(Error)
      expect((failure as Error).message).toContain("Sign-in completed, but a token could still not be obtained")
      expect((failure as Error).cause).toBeInstanceOf(AuthenticationRequiredError)
    })

    it("does not turn other failures into a sign-in", async () => {
      const error = new Error("network down")
      vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockRejectedValue(error)
      const authenticate = vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate")

      await expect((createCredential(config).value as TokenCredential).getToken("scope")).rejects.toBe(error)
      expect(authenticate).not.toHaveBeenCalled()
    })

    // Without the record, the device-code credential has no account to look up in the
    // cache, so a persisted refresh token would never be used.
    describe("with device code", () => {
      const deviceConfig = { ...config, useDeviceCode: true }

      it("signs in through authenticate() and persists the record it returns", async () => {
        const getToken = vi
          .spyOn(DeviceCodeCredential.prototype, "getToken")
          .mockRejectedValueOnce(required())
          .mockResolvedValue(token)
        const authenticate = vi.spyOn(DeviceCodeCredential.prototype, "authenticate").mockResolvedValue(record)
        const browserToken = vi.spyOn(InteractiveBrowserCredential.prototype, "getToken")
        const credential = createCredential(deviceConfig).value as TokenCredential

        await expect(credential.getToken("scope")).resolves.toBe(token)

        expect(authenticate).toHaveBeenCalledOnce()
        expect(getToken).toHaveBeenCalledTimes(2)
        expect(browserToken).not.toHaveBeenCalled()
        expect(readAuthenticationRecord(cacheDir.path)).toEqual(record)
      })

      it("does not prompt when the cache serves the token", async () => {
        vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
        const authenticate = vi.spyOn(DeviceCodeCredential.prototype, "authenticate")
        const credential = createCredential(deviceConfig).value as TokenCredential

        await expect(credential.getToken("scope")).resolves.toBe(token)
        expect(authenticate).not.toHaveBeenCalled()
      })

      it("picks up the record saved by an earlier start", async () => {
        writeAuthenticationRecord(record, cacheDir.path)
        const log = vi.spyOn(console, "error").mockImplementation(() => undefined)
        vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockRejectedValueOnce(required()).mockResolvedValue(token)
        vi.spyOn(DeviceCodeCredential.prototype, "authenticate").mockResolvedValue(record)

        await (createCredential(deviceConfig).value as TokenCredential).getToken("scope")
        expect(log).toHaveBeenCalledWith(expect.stringContaining("Cached credentials could not be reused"))
      })
    })

    // An unattended host has nobody to enter a device code; waiting for one only stalls it.
    describe("with sign-in disabled", () => {
      const unattended = { ...config, useDeviceCode: true, disableSignIn: true }

      it("fails at once instead of prompting", async () => {
        writeAuthenticationRecord(record, cacheDir.path)
        vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockRejectedValue(required())
        const authenticate = vi.spyOn(DeviceCodeCredential.prototype, "authenticate")
        const credential = createCredential(unattended).value as TokenCredential

        const failure = await credential.getToken("scope").catch((e: unknown) => e)
        expect(failure).toBeInstanceOf(Error)
        expect((failure as Error).message).toContain("the saved sign-in could not be reused")
        expect((failure as Error).message).toContain("MS365_DISABLE_SIGN_IN")
        // A saved sign-in can fail for a passing reason (the network); the Azure log says which.
        expect((failure as Error).message).toContain("AZURE_LOG_LEVEL=info")
        expect((failure as Error).cause).toBeInstanceOf(AuthenticationRequiredError)
        expect(authenticate).not.toHaveBeenCalled()
      })

      it("applies to the browser flow too", async () => {
        vi.spyOn(InteractiveBrowserCredential.prototype, "getToken").mockRejectedValue(required())
        const authenticate = vi.spyOn(InteractiveBrowserCredential.prototype, "authenticate")
        const credential = createCredential({ ...config, disableSignIn: true }).value as TokenCredential

        const failure = await credential.getToken("scope").catch((e: unknown) => e)
        expect((failure as Error).message).toContain("there is no saved sign-in")
        // With nothing saved, there is nothing for the log to explain.
        expect((failure as Error).message).not.toContain("AZURE_LOG_LEVEL")
        expect(authenticate).not.toHaveBeenCalled()
      })

      it("still serves tokens from the cache", async () => {
        vi.spyOn(DeviceCodeCredential.prototype, "getToken").mockResolvedValue(token)
        const credential = createCredential(unattended).value as TokenCredential

        await expect(credential.getToken("scope")).resolves.toBe(token)
      })

      it("is refused when the token cache is off, since every start would need a sign-in", () => {
        vi.stubEnv("MS365_TOKEN_CACHE", "false")
        const result = createCredential(unattended)
        expect(result.isLeft()).toBe(true)
        expect((result.value as { message: string }).message).toContain("MS365_TOKEN_CACHE=false")
      })
    })

    it("ignores a stored record missing its account fields", () => {
      writeAuthenticationRecord({ ...record, homeAccountId: "" }, cacheDir.path)
      expect(readAuthenticationRecord(cacheDir.path)).toBeUndefined()
    })
  })
})
