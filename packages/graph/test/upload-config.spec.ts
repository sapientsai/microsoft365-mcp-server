import { isUploadTicket, resolveUploadTicket } from "@sapientsai/ms-graph-core"
import { describe, expect, it } from "vitest"

import { buildGetUploadConfigTool } from "../src/tools/upload-config"

const PATH = "/me/drive/root:/Documents/report.docx:/content"

describe("get_upload_config tool", () => {
  it("never echoes the raw api key — hands back an opaque ticket that resolves to it", async () => {
    const tool = buildGetUploadConfigTool("https://graph.example.com", "RAW_SECRET_KEY")
    const out = await tool.execute({ path: PATH, conflict_behavior: "rename" })

    expect(out).not.toContain("RAW_SECRET_KEY")
    const payload = JSON.parse(out) as { authHeader: string; uploadUrl: string; contentType: string }
    const ticket = payload.authHeader.replace(/^Authorization: Bearer\s+/, "")
    expect(isUploadTicket(ticket)).toBe(true)
    expect(resolveUploadTicket(ticket)).toBe("RAW_SECRET_KEY")
    expect(payload.uploadUrl).toContain("https://graph.example.com/upload?")
    expect(payload.contentType).toContain("wordprocessingml") // .docx inferred
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
