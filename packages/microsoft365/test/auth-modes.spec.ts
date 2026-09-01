import type { AccessToken, TokenCredential } from "@azure/identity"
import { beforeEach, describe, expect, it, vi } from "vitest"

import { isBrowserLaunchFailure, withDeviceCodeFallback } from "../src/auth/auth-modes"

const TOKEN: AccessToken = { token: "t", expiresOnTimestamp: 0 }

const credential = (impl: () => Promise<AccessToken | null>): TokenCredential => ({ getToken: impl })

beforeEach(() => {
  vi.spyOn(console, "error").mockImplementation(() => {})
})

describe("isBrowserLaunchFailure", () => {
  it.each([
    "Arc is already open. Only one instance of Arc can be opened at a time.",
    "Unable to open browser",
    "spawn open ENOENT",
    "no such file or directory",
  ])("should treat %s as a launch failure", (message) => {
    expect(isBrowserLaunchFailure(new Error(message))).toBe(true)
  })

  it("should match case-insensitively", () => {
    expect(isBrowserLaunchFailure(new Error("ONLY ONE INSTANCE"))).toBe(true)
  })

  it("should handle a non-Error value", () => {
    expect(isBrowserLaunchFailure("could not open browser")).toBe(true)
    expect(isBrowserLaunchFailure(undefined)).toBe(false)
  })

  it.each([
    "AADSTS65004: User declined to consent",
    "invalid_client: client id is malformed",
    "getaddrinfo ENOTFOUND login.microsoftonline.com",
  ])("should not treat %s as a launch failure", (message) => {
    expect(isBrowserLaunchFailure(new Error(message))).toBe(false)
  })
})

describe("withDeviceCodeFallback", () => {
  it("should return the browser token when the browser works", async () => {
    const fallback = vi.fn()
    const wrapped = withDeviceCodeFallback(
      credential(() => Promise.resolve(TOKEN)),
      "tenant",
      "client",
      fallback,
    )

    await expect(wrapped.getToken("scope")).resolves.toBe(TOKEN)
    expect(fallback).not.toHaveBeenCalled()
  })

  it("should fall back to device code when the browser cannot launch", async () => {
    const deviceToken: AccessToken = { token: "device", expiresOnTimestamp: 0 }
    const fallback = vi.fn(() => credential(() => Promise.resolve(deviceToken)))
    const wrapped = withDeviceCodeFallback(
      credential(() => Promise.reject(new Error("Arc is already open. Only one instance"))),
      "tenant",
      "client",
      fallback,
    )

    await expect(wrapped.getToken("scope")).resolves.toBe(deviceToken)
    expect(fallback).toHaveBeenCalledWith("tenant", "client")
  })

  it("should rethrow a non-launch error rather than hide it behind a second prompt", async () => {
    const fallback = vi.fn()
    const wrapped = withDeviceCodeFallback(
      credential(() => Promise.reject(new Error("AADSTS65004: User declined to consent"))),
      "tenant",
      "client",
      fallback,
    )

    await expect(wrapped.getToken("scope")).rejects.toThrow("declined to consent")
    expect(fallback).not.toHaveBeenCalled()
  })

  it("should pass scopes and options through to both credentials", async () => {
    const browser = vi.fn(() => Promise.reject(new Error("unable to open browser")))
    const device = vi.fn(() => Promise.resolve(TOKEN))
    const wrapped = withDeviceCodeFallback(credential(browser), "tenant", "client", () => credential(device))

    const options = { requestOptions: { timeout: 1 } }
    await wrapped.getToken(["a", "b"], options)

    expect(browser).toHaveBeenCalledWith(["a", "b"], options)
    expect(device).toHaveBeenCalledWith(["a", "b"], options)
  })
})
