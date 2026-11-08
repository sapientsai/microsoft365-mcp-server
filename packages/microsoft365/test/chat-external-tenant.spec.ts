import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({ getGraphClient: vi.fn() }))
vi.mock("../src/auth", () => ({ getAccessToken: vi.fn(), getAuthStatus: vi.fn() }))

import { getAccessToken } from "../src/auth"
import { getGraphClient } from "../src/client/graph-client"
import { listChatMessages } from "../src/tools/chat-tools"

// Made-up tenants.
const OWN = "11111111-1111-1111-1111-111111111111"
const OTHER = "22222222-2222-2222-2222-222222222222"
const token = (tid: string) =>
  `h.${Buffer.from(JSON.stringify({ tid, scp: "Chat.ReadWrite" })).toString("base64url")}.s`

const mockClient = { getMe: vi.fn(), listChatMessages: vi.fn(), request: vi.fn(), requestPaginated: vi.fn() }
const insufficient = { type: "forbidden", status: 403, message: "InsufficientPrivileges" }

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
  vi.mocked(getAccessToken).mockResolvedValue(Right(token(OWN)) as never)
  mockClient.getMe.mockResolvedValue(Right({ id: "me" }))
  mockClient.listChatMessages.mockResolvedValue(Left(insufficient))
  mockClient.requestPaginated.mockResolvedValue(Left(insufficient))
})

describe("list_chat_messages on a chat another organization hosts", () => {
  it("names the other organization when the chat's tenant differs from the token's", async () => {
    mockClient.request.mockResolvedValue(Right({ tenantId: OTHER }))
    const result = await listChatMessages({ chat_id: "19:meeting_abc@thread.v2" })
    const message = (result.value as Error).message
    expect(message).toContain("Failed to list chat messages: InsufficientPrivileges")
    expect(message).toContain(`belongs to another organization (tenant ${OTHER})`)
    expect(mockClient.request).toHaveBeenCalledWith("GET", "/chats/19:meeting_abc@thread.v2", {
      odataParams: { $select: ["tenantId"] },
    })
  })

  it("does the same when every page is asked for", async () => {
    mockClient.request.mockResolvedValue(Right({ tenantId: OTHER }))
    const result = await listChatMessages({ chat_id: "c1", fetch_all_pages: true })
    expect((result.value as Error).message).toContain("another organization")
  })

  it("keeps Graph's words for a chat in the token's own tenant", async () => {
    mockClient.request.mockResolvedValue(Right({ tenantId: OWN.toUpperCase() }))
    const result = await listChatMessages({ chat_id: "c1" })
    expect((result.value as Error).message).toBe("Failed to list chat messages: InsufficientPrivileges")
  })

  it("keeps Graph's words when the chat can't be read either", async () => {
    mockClient.request.mockResolvedValue(Left(insufficient))
    const result = await listChatMessages({ chat_id: "c1" })
    expect((result.value as Error).message).toBe("Failed to list chat messages: InsufficientPrivileges")
  })

  it("doesn't read the chat for a failure that isn't a refusal", async () => {
    mockClient.listChatMessages.mockResolvedValue(Left({ type: "network", message: "offline" }))
    const result = await listChatMessages({ chat_id: "c1" })
    expect((result.value as Error).message).toBe("Failed to list chat messages: offline")
    expect(mockClient.request).not.toHaveBeenCalled()
  })
})
