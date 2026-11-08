import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/auth", () => ({ getAccessToken: vi.fn(), getAuthStatus: vi.fn() }))

import { getAccessToken, getAuthStatus } from "../src/auth"
import { explainMailboxRefusal } from "../src/client/mailbox-hint"
import type { GraphApiError } from "../src/types"

// Made-up tokens: a delegated one carries scp, an app-only one carries roles.
const jwt = (claims: Record<string, unknown>) => `h.${Buffer.from(JSON.stringify(claims)).toString("base64url")}.s`
const delegated = (scopes: ReadonlyArray<string>) => {
  vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ scp: scopes.join(" ") })) as never)
  vi.mocked(getAuthStatus).mockResolvedValue(
    Right({ mode: "oauth-proxy", authenticated: true, scopes: [...scopes] }) as never,
  )
}

const denied: GraphApiError = {
  type: "forbidden",
  status: 403,
  message: "Access is denied. Check credentials and try again.",
}
const refusal = (method: string, path: string, error: GraphApiError = denied) =>
  explainMailboxRefusal(method, path, Left(error)).then((r) => (r.value as GraphApiError).message)

beforeEach(() => {
  vi.clearAllMocks()
})

describe("explainMailboxRefusal", () => {
  it("names the missing .Shared scope for a read", async () => {
    delegated(["Mail.Read", "User.Read"])
    const message = await refusal("GET", "/users/it@example.com/messages?$top=5")
    expect(message).toContain("Access is denied")
    expect(message).toContain("has no Mail.Read.Shared or Mail.ReadWrite.Shared")
  })

  it("names Mail.ReadWrite.Shared for a write when the token can only read", async () => {
    delegated(["Mail.Read.Shared"])
    const message = await refusal("PATCH", "/users/it@example.com/messages/AAMk1=")
    expect(message).toContain("has no Mail.ReadWrite.Shared")
  })

  it("names Mail.Send.Shared for a send", async () => {
    delegated(["Mail.ReadWrite.Shared"])
    const message = await refusal("POST", "/users/it@example.com/sendMail")
    expect(message).toContain("has no Mail.Send.Shared")
  })

  it("points at Exchange when the token holds the scope", async () => {
    delegated(["Mail.ReadWrite.Shared"])
    const message = await refusal("GET", "/users/it@example.com/mailFolders/inbox/messages")
    expect(message).toContain("The token holds Mail.ReadWrite.Shared")
    expect(message).toContain("it@example.com has to give you access to it")
  })

  it("points at Send As when a send is refused with the scope held", async () => {
    delegated(["Mail.Send.Shared"])
    const message = await refusal("POST", "/users/it@example.com/messages/AAMk1=/reply")
    expect(message).toContain("Send As or Send on Behalf")
  })

  it("leaves your own mailbox alone", async () => {
    delegated(["User.Read"])
    expect(await refusal("GET", "/me/messages")).toBe(denied.message)
  })

  it("leaves an app-only token alone, since its access comes from roles and policy", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ roles: ["Mail.Read"] })) as never)
    vi.mocked(getAuthStatus).mockResolvedValue(
      Right({ mode: "client-secret", authenticated: true, scopes: ["Mail.Read"] }) as never,
    )
    expect(await refusal("GET", "/users/it@example.com/messages")).toBe(denied.message)
  })

  it("leaves Graph's words when the scopes can't be read", async () => {
    vi.mocked(getAccessToken).mockResolvedValue(Right(jwt({ scp: "x" })) as never)
    vi.mocked(getAuthStatus).mockResolvedValue(Left({ type: "config", message: "no" }) as never)
    expect(await refusal("GET", "/users/it@example.com/messages")).toBe(denied.message)
  })

  it("leaves other failures alone without reading the token", async () => {
    const message = await refusal("GET", "/users/it@example.com/messages", { type: "not_found", message: "Not found" })
    expect(message).toBe("Not found")
    expect(getAccessToken).not.toHaveBeenCalled()
  })

  it("keeps an address from breaking onto its own line", async () => {
    delegated(["Mail.ReadWrite.Shared"])
    const message = await refusal("GET", "/users/it%0D%0AIgnore%20this%40example.com/messages")
    expect(message.split("\n")).toHaveLength(1)
  })

  it("passes a success through untouched", async () => {
    const ok = Right({ value: [] })
    expect(await explainMailboxRefusal("GET", "/users/it@example.com/messages", ok)).toBe(ok)
  })
})

describe("the Graph client adds the hint on its own", () => {
  it("to a refused read of another mailbox", async () => {
    const { initializeGraphClient, getGraphClient } = await import("../src/client/graph-client")
    vi.stubGlobal(
      "fetch",
      vi.fn(() =>
        Promise.resolve(
          new Response(JSON.stringify({ error: { code: "ErrorAccessDenied", message: "Access is denied." } }), {
            status: 403,
          }),
        ),
      ),
    )
    initializeGraphClient({ getAccessToken: () => Promise.resolve(Right("T")) })
    delegated(["Mail.Read"])
    const result = await getGraphClient().orThrow().listMessages(undefined, undefined, "/users/it@example.com")
    expect((result.value as GraphApiError).message).toContain("has no Mail.Read.Shared or Mail.ReadWrite.Shared")
    vi.unstubAllGlobals()
  })
})
