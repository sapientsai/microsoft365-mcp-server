import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))

import { mintUploadTicket } from "@sapientsai/ms-graph-core"

const TARGET = { relay: "drive", path: "/me/drive/root:/a.txt:/content" } as const

import { getAccessToken } from "../src/auth"
import { resolveUploadAccessToken } from "../src/upload/upload-auth"

describe("resolveUploadAccessToken", () => {
  beforeEach(() => {
    vi.clearAllMocks()
    delete process.env.MS365_UPLOAD_TOKEN
  })
  afterEach(() => {
    delete process.env.MS365_UPLOAD_TOKEN
  })

  it("resolves an opaque ticket to the server-held token (any mode), once, for its upload", async () => {
    const ticket = mintUploadTicket("GRAPH.TOKEN", TARGET)
    expect(await resolveUploadAccessToken(false, ticket, TARGET)).toEqual({ token: "GRAPH.TOKEN" })
    expect(getAccessToken).not.toHaveBeenCalled()
    const again = await resolveUploadAccessToken(true, ticket, TARGET)
    expect(again.status).toBe(401)
    expect(again.error).toContain("already used")
  })

  it("still accepts a shared secret that happens to start like a ticket, every time", async () => {
    process.env.MS365_UPLOAD_TOKEN = "upl_made-up-secret"
    vi.mocked(getAccessToken).mockResolvedValue(Right("SERVER.TOKEN"))
    expect(await resolveUploadAccessToken(false, "upl_made-up-secret", TARGET)).toEqual({ token: "SERVER.TOKEN" })
    expect(await resolveUploadAccessToken(false, "upl_made-up-secret", TARGET)).toEqual({ token: "SERVER.TOKEN" })
  })

  // In oauth-proxy mode a non-ticket bearer goes to Graph as the token; a spent or misdirected
  // ticket must not, so it is refused rather than passed through.
  it("refuses a ticket issued for another upload, in either mode", async () => {
    const other = { relay: "drive", path: "/me/drive/root:/other.txt:/content" } as const
    for (const oauthMode of [true, false]) {
      const result = await resolveUploadAccessToken(oauthMode, mintUploadTicket("GRAPH.TOKEN", TARGET), other)
      expect(result.status).toBe(401)
      expect(result.token).toBeUndefined()
    }
  })

  describe("oauth-proxy mode", () => {
    it("requires a bearer", async () => {
      expect(await resolveUploadAccessToken(true, undefined, TARGET)).toEqual({
        error: "Missing Bearer token",
        status: 401,
      })
    })
    it("passes the delegated bearer through", async () => {
      expect(await resolveUploadAccessToken(true, "delegated-jwt", TARGET)).toEqual({ token: "delegated-jwt" })
    })
  })

  describe("shared-secret (non-oauth) mode", () => {
    it("REFUSES (503) when no MS365_UPLOAD_TOKEN is configured — no unauthenticated server-credential upload", async () => {
      const result = await resolveUploadAccessToken(false, "anything", TARGET)
      expect(result.status).toBe(503)
      expect(getAccessToken).not.toHaveBeenCalled()
    })

    it("rejects a wrong token with 401", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      expect(await resolveUploadAccessToken(false, "wrong", TARGET)).toEqual({
        error: "Invalid upload token",
        status: 401,
      })
      expect(getAccessToken).not.toHaveBeenCalled()
    })

    it("accepts the correct token and returns the server token", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      vi.mocked(getAccessToken).mockResolvedValue(Right("SERVER.TOKEN") as never)
      expect(await resolveUploadAccessToken(false, "s3cret", TARGET)).toEqual({ token: "SERVER.TOKEN" })
    })

    it("surfaces a server-side token failure as 401", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      vi.mocked(getAccessToken).mockResolvedValue(Left({ message: "no creds" }) as never)
      expect(await resolveUploadAccessToken(false, "s3cret", TARGET)).toEqual({ error: "no creds", status: 401 })
    })
  })
})
