import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth/token-context", () => ({ getContextToken: vi.fn() }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))

import { isUploadTicket, resolveUploadTicket } from "@sapientsai/ms-graph-core"

import { getAccessToken } from "../src/auth"
import { getContextToken } from "../src/auth/token-context"
import { getUploadConfig } from "../src/tools/files-tools"

// upload/upload-ticket/describeFetchError unit tests live in @sapientsai/ms-graph-core.
// This exercises getUploadConfig's token-hardening (server-side glue over core).
describe("getUploadConfig token hardening", () => {
  const RAW_JWT = "eyJ.DELEGATED-GRAPH-JWT.signature"

  beforeEach(() => {
    vi.clearAllMocks()
    process.env.MS365_PUBLIC_BASE_URL = "https://ms365.example.com"
    vi.mocked(getContextToken).mockReturnValue(RAW_JWT)
    vi.stubGlobal(
      "fetch",
      vi.fn(() => Promise.resolve({ status: 401 } as Response)),
    )
  })

  afterEach(() => {
    vi.unstubAllGlobals()
    delete process.env.MS365_PUBLIC_BASE_URL
  })

  it("never echoes the raw Graph token; returns an opaque ticket that resolves to it", async () => {
    const result = await getUploadConfig({ path: "/me/drive/root:/_upload-test/probe.bin:/content" })
    expect(result.isRight()).toBe(true)
    const json = result.value as string

    // The raw delegated JWT must not appear anywhere in the client-facing payload.
    expect(json).not.toContain(RAW_JWT)

    const payload = JSON.parse(json) as {
      authHeader: string
      graphReachableFromServer: boolean
      graphReachabilityDetail: string
    }
    const ticket = payload.authHeader.replace(/^Authorization: Bearer\s+/, "")
    expect(isUploadTicket(ticket)).toBe(true)
    expect(resolveUploadTicket(ticket)).toBe(RAW_JWT)
    expect(payload.graphReachableFromServer).toBe(true)
    expect(getAccessToken).not.toHaveBeenCalled()
  })
})
