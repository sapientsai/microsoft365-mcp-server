// The upload relay is called with a curl command that prints its response, and the upload config
// tools print the header that curl sends. Neither may carry a credential that outlives the upload:
// Graph's pre-authenticated download URL, or the shared MS365_UPLOAD_TOKEN, which never expires.
import { Left, Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

const auth = vi.hoisted(() => ({ contextToken: undefined as string | undefined }))
vi.mock("../src/auth/token-context", () => ({ getContextToken: vi.fn(() => auth.contextToken) }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn() }))

import { withoutDownloadUrl } from "@sapientsai/ms-graph-core"

import { getAccessToken } from "../src/auth"
import { relayAuthHeader } from "../src/tools/files-tools"

const SHARED_SECRET = "made-up-shared-upload-secret"
const DOWNLOAD_URL = "https://example.sharepoint.com/download.aspx?tempauth=made-up-tempauth"

beforeEach(() => {
  auth.contextToken = undefined
  vi.stubEnv("MS365_UPLOAD_TOKEN", SHARED_SECRET)
})
afterEach(() => vi.unstubAllEnvs())

describe("relayAuthHeader", () => {
  it("hands out a ticket, never the Graph token, when one resolves", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Right("GRAPH.TOKEN.VALUE"))
    const header = await relayAuthHeader()
    expect(header).toMatch(/^Authorization: Bearer /)
    expect(header).not.toContain("GRAPH.TOKEN.VALUE")
    expect(header).not.toContain(SHARED_SECRET)
  })

  it("hands out nothing, not the shared secret, when no Graph token resolves", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Left({ type: "auth", message: "not signed in" }))
    expect(await relayAuthHeader()).toBeUndefined()
  })
})

describe("withoutDownloadUrl", () => {
  it("drops both forms of the download URL and keeps the rest of the item", () => {
    const item = {
      id: "ITEM1",
      name: "a.txt",
      eTag: '"{A},1"',
      "@microsoft.graph.downloadUrl": DOWNLOAD_URL,
      "@content.downloadUrl": DOWNLOAD_URL,
    }
    const cleaned = withoutDownloadUrl(item)
    expect(cleaned).toEqual({ id: "ITEM1", name: "a.txt", eTag: '"{A},1"' })
    expect(JSON.stringify(cleaned)).not.toContain("tempauth")
  })
})
