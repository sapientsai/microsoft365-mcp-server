// Checks for caller values that end up in a request header. A bad one is refused with a message
// that names the parameter, before any request is built. The message never echoes the value.

/** Visible ASCII and space, nothing else: what an eTag, MIME type or bearer token is made of. */
export const isPrintableAscii = (value: string): boolean => /^[\x20-\x7e]+$/.test(value)

export const blankIfMatchMessage = (label: string): string =>
  `${label} is blank. Pass the eTag from download_file, get_drive_item or list_drive_items, or leave it out.`

export const BLANK_IF_MATCH_MESSAGE = blankIfMatchMessage("if_match")

/**
 * Why an If-Match value can't be used, or undefined when it can (or was left out). A blank one
 * would go out as an empty header, which Graph may ignore, so the write would run unconditionally.
 */
export const ifMatchProblem = (ifMatch: string | undefined, label = "if_match"): string | undefined => {
  if (ifMatch === undefined) return undefined
  if (ifMatch.trim() === "") return blankIfMatchMessage(label)
  if (!isPrintableAscii(ifMatch)) {
    return `${label} holds characters an eTag never has, such as a line break. Pass the eTag exactly as it was printed.`
  }
  return undefined
}

/** Why a content type can't be sent, or undefined when it can (or was left out). */
export const contentTypeProblem = (contentType: string | undefined, label = "content_type"): string | undefined => {
  if (contentType === undefined) return undefined
  if (contentType.trim() === "" || !isPrintableAscii(contentType)) {
    return `${label} must be a MIME type such as text/plain or application/pdf, with no line breaks or other special characters.`
  }
  return undefined
}

export const CONFLICT_BEHAVIORS = ["rename", "replace", "fail"] as const

/** Why a conflict behavior can't be used, or undefined when it can (or was left out). */
export const conflictBehaviorProblem = (value: string | undefined, label = "conflictBehavior"): string | undefined =>
  value === undefined || (CONFLICT_BEHAVIORS as ReadonlyArray<string>).includes(value)
    ? undefined
    : `${label} must be one of ${CONFLICT_BEHAVIORS.join(", ")}.`
