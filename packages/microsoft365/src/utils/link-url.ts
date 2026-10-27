import type { GraphAttachment } from "../types"

// A referenceAttachment's sourceUrl is chosen by whoever sent the email, the same as its name, so
// it is never printed as received. It is shown only when it parses as an http(s) URL, and then as
// the URL parser serialises it: tabs and newlines are removed and other control and non-ASCII
// characters percent-encoded, so it cannot break out of its line or forge another attachment row.
export const displayableLinkUrl = (raw: string | undefined): string | undefined => {
  if (raw === undefined || !URL.canParse(raw)) return undefined
  const url = new URL(raw)
  return url.protocol === "https:" || url.protocol === "http:" ? url.href : undefined
}

/** Where a cloud-link attachment points, worded as data rather than as an instruction. */
export const linkLocation = (att: GraphAttachment): string => {
  if (att.sourceUrl === undefined) return "link not returned by Graph"
  return displayableLinkUrl(att.sourceUrl) ?? "link withheld: not an http(s) URL"
}
