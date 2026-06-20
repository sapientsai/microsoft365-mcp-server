import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))

import { getAccessToken } from "../src/auth"
import { resolveUploadAccessToken } from "../src/upload/upload-auth"
import { mintUploadTicket } from "../src/upload/upload-ticket"

describe("resolveUploadAccessToken", () => {
  beforeEach(() => {
    vi.clearAllMocks()
    delete process.env.MS365_UPLOAD_TOKEN
  })
  afterEach(() => {
    delete process.env.MS365_UPLOAD_TOKEN
  })

  it("resolves an opaque ticket to the server-held token (any mode)", async () => {
    const ticket = mintUploadTicket("GRAPH.TOKEN")
    const result = await resolveUploadAccessToken(false, ticket)
    expect(result).toEqual({ token: "GRAPH.TOKEN" })
    expect(getAccessToken).not.toHaveBeenCalled()
  })

  describe("oauth-proxy mode", () => {
    it("requires a bearer", async () => {
      expect(await resolveUploadAccessToken(true, undefined)).toEqual({ error: "Missing Bearer token", status: 401 })
    })
    it("passes the delegated bearer through", async () => {
      expect(await resolveUploadAccessToken(true, "delegated-jwt")).toEqual({ token: "delegated-jwt" })
    })
  })

  describe("shared-secret (non-oauth) mode", () => {
    it("REFUSES (503) when no MS365_UPLOAD_TOKEN is configured — no unauthenticated server-credential upload", async () => {
      const result = await resolveUploadAccessToken(false, "anything")
      expect(result.status).toBe(503)
      expect(getAccessToken).not.toHaveBeenCalled()
    })

    it("rejects a wrong token with 401", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      expect(await resolveUploadAccessToken(false, "wrong")).toEqual({ error: "Invalid upload token", status: 401 })
      expect(getAccessToken).not.toHaveBeenCalled()
    })

    it("accepts the correct token and returns the server token", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      vi.mocked(getAccessToken).mockResolvedValue(Right("SERVER.TOKEN") as never)
      expect(await resolveUploadAccessToken(false, "s3cret")).toEqual({ token: "SERVER.TOKEN" })
    })

    it("surfaces a server-side token failure as 401", async () => {
      process.env.MS365_UPLOAD_TOKEN = "s3cret"
      vi.mocked(getAccessToken).mockResolvedValue(Left({ message: "no creds" }) as never)
      expect(await resolveUploadAccessToken(false, "s3cret")).toEqual({ error: "no creds", status: 401 })
    })
  })
})
