import { mintUploadTicket } from "@sapientsai/ms-graph-core"
import { describe, expect, it } from "vitest"

import { authorizesWithApiKey } from "../src/auth/api-key-gate"

describe("authorizesWithApiKey", () => {
  it("accepts the raw api key", () => {
    expect(authorizesWithApiKey("SECRET", "SECRET")).toBe(true)
  })

  it("accepts an opaque upload ticket that resolves to the api key", () => {
    const ticket = mintUploadTicket("SECRET")
    expect(authorizesWithApiKey(ticket, "SECRET")).toBe(true)
  })

  it("rejects a wrong key", () => {
    expect(authorizesWithApiKey("wrong", "SECRET")).toBe(false)
  })

  it("rejects a missing bearer", () => {
    expect(authorizesWithApiKey(undefined, "SECRET")).toBe(false)
  })

  it("rejects an unknown ticket-shaped value", () => {
    expect(authorizesWithApiKey("upl_neverminted", "SECRET")).toBe(false)
  })
})
