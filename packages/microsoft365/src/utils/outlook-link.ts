// A task in the Flagged email list links to its email with an Outlook web URL, not a Graph id: its
// linkedResources carry applicationName "Outlook", an externalId no mail tool uses, and a webUrl
// such as https://outlook.office365.com/owa/?ItemID=<id>&… whose ItemID is the message's id in
// Outlook's older format. Swapping "/" for "-" and "+" for "_" gives the Graph id; a GET of that id
// returned the flagged message (live, 2026-10-30). Microsoft doesn't document the mapping, so the
// result is checked, and anything that doesn't fit is left out rather than guessed.
import type { GraphLinkedResource } from "../types"
import { isGraphId } from "./graph-id"

const OUTLOOK_HOSTS = new Set(["outlook.office365.com", "outlook.office.com", "outlook.live.com"])

export const linkedMessageId = (link: GraphLinkedResource): string | undefined => {
  if (link.applicationName?.toLowerCase() !== "outlook" || link.webUrl === undefined) return undefined
  if (!URL.canParse(link.webUrl)) return undefined
  const url = new URL(link.webUrl)
  if (url.protocol !== "https:" || !OUTLOOK_HOSTS.has(url.hostname)) return undefined
  const id = url.searchParams.get("ItemID")?.replace(/\//g, "-").replace(/\+/g, "_")
  return id !== undefined && isGraphId(id) ? id : undefined
}
