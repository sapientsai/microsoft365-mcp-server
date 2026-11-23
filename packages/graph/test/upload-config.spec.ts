import { isUploadTicket, mintUploadTicket, redeemUploadTicket } from "@sapientsai/ms-graph-core"
import { describe, expect, it, vi } from "vitest"

// The real minting, wrapped so a test can see whether a ticket was minted at all.
vi.mock("@sapientsai/ms-graph-core", async (original) => {
  const real = await original<typeof import("@sapientsai/ms-graph-core")>()
  return { ...real, mintUploadTicket: vi.fn(real.mintUploadTicket) }
})

import { buildGetUploadConfigTool, NO_PUBLIC_BASE_URL_MESSAGE } from "../src/tools/upload-config"

const PATH = "/me/drive/root:/Documents/report.docx:/content"

describe("get_upload_config tool", () => {
  it("never echoes the raw api key — hands back a ticket good only for this path", async () => {
    const tool = buildGetUploadConfigTool("https://graph.example.com", "RAW_SECRET_KEY")
    const out = await tool.execute({ path: PATH, conflict_behavior: "rename" })

    expect(out).not.toContain("RAW_SECRET_KEY")
    const payload = JSON.parse(out) as { authHeader: string; uploadUrl: string; contentType: string }
    const ticket = payload.authHeader.replace(/^Authorization: Bearer\s+/, "")
    expect(isUploadTicket(ticket)).toBe(true)
    expect(redeemUploadTicket(ticket, { relay: "drive", path: "/me/drive/root:/other.docx:/content" })).toBeUndefined()
    const again = (
      JSON.parse(await tool.execute({ path: PATH, conflict_behavior: "rename" })) as { authHeader: string }
    ).authHeader.replace(/^Authorization: Bearer\s+/, "")
    expect(redeemUploadTicket(again, { relay: "drive", path: PATH })).toBe("RAW_SECRET_KEY")
    expect(payload.uploadUrl).toContain("https://graph.example.com/upload?")
    expect(payload.contentType).toContain("wordprocessingml") // .docx inferred
  })

  it("refuses, before minting a ticket, when it has no public URL to give", async () => {
    vi.mocked(mintUploadTicket).mockClear()
    const tool = buildGetUploadConfigTool(undefined, "RAW_SECRET_KEY")
    await expect(tool.execute({ path: PATH, conflict_behavior: "rename" })).rejects.toThrow(NO_PUBLIC_BASE_URL_MESSAGE)
    expect(mintUploadTicket).not.toHaveBeenCalled()
  })

  it("omits the auth header and warns when no api key is configured", async () => {
    const out = await buildGetUploadConfigTool("https://graph.example.com", undefined).execute({
      path: PATH,
      conflict_behavior: "rename",
    })
    const payload = JSON.parse(out) as { authHeader?: string; notes: string[] }
    expect(payload.authHeader).toBeUndefined()
    expect(payload.notes.some((n) => n.includes("503"))).toBe(true)
  })

  it("rejects a path that does not end in :/content", async () => {
    await expect(
      buildGetUploadConfigTool("https://x", "k").execute({
        path: "/me/drive/root:/x.docx",
        conflict_behavior: "rename",
      }),
    ).rejects.toThrow(":/content")
  })
})
