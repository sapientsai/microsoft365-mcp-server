// A link URL on a task or attachment was chosen by whoever created the item, so it is shown only
// when it parses as http: or https:, and then as the parser's serialisation (url.href), which
// percent-encodes control characters and can't carry a newline into the output.
export const displayableLinkUrl = (raw: string | undefined): string | undefined => {
  if (raw === undefined || !URL.canParse(raw)) return undefined
  const url = new URL(raw)
  return url.protocol === "https:" || url.protocol === "http:" ? url.href : undefined
}
