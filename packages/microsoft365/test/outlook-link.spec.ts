import { describe, expect, it } from "vitest"

import { linkedMessageId } from "../src/utils/outlook-link"

const OWA = "https://outlook.office365.com/owa/?ItemID=AAMk%2FAB%2Bc%3D&exvsurl=1&viewmodel=ReadMessageItem"

describe("linkedMessageId", () => {
  it("turns an Outlook link's ItemID into a Graph message id", () => {
    expect(linkedMessageId({ applicationName: "Outlook", webUrl: OWA })).toBe("AAMk-AB_c=")
  })

  it.each([
    ["another app", { applicationName: "Claude test", webUrl: OWA }],
    ["no link", { applicationName: "Outlook" }],
    ["a link that doesn't parse", { applicationName: "Outlook", webUrl: "not a url" }],
    ["another host", { applicationName: "Outlook", webUrl: "https://evil.example/owa/?ItemID=AAMk" }],
    ["plain http", { applicationName: "Outlook", webUrl: "http://outlook.office365.com/owa/?ItemID=AAMk" }],
    ["no ItemID", { applicationName: "Outlook", webUrl: "https://outlook.office365.com/owa/" }],
    [
      "an ItemID with characters an id can't hold",
      { applicationName: "Outlook", webUrl: "https://outlook.office365.com/owa/?ItemID=a)%20(ID%3A%20x" },
    ],
  ])("gives nothing for %s", (_label, link) => {
    expect(linkedMessageId(link)).toBe(undefined)
  })
})
