import type { Either } from "functype/either"
import { Left } from "functype/either"

import { heldScope, tokenKind, tokenScopes } from "../auth/token-facts"
import type { GraphApiError } from "../types"
import { safeDecode } from "../utils/safe-decode"

// Graph refuses another person's mailbox with the same "Access is denied" whether the token lacks the
// .Shared scope or Exchange never gave the user access to that mailbox, and the two have different
// fixes. The token's scopes tell them apart, so a refusal on /users/{address}/ mail says which it is.
// Delegated tokens only: an app-only token reaches mailboxes through app roles and an access policy,
// and nothing here can see either.
const OTHER_MAILBOX = /^\/users\/([^/?()]+)\/(?:messages|mailFolders|sendMail)(?=[/?(]|$)/i
const SENDS = /\/(?:sendMail|send|reply|replyAll|forward)$/i

const READ = ["Mail.Read.Shared", "Mail.ReadWrite.Shared"] as const
const WRITE = ["Mail.ReadWrite.Shared"] as const
const SEND = ["Mail.Send.Shared"] as const

const oneLine = (value: string): string => value.replace(/[\r\n]+/g, " ").slice(0, 120)

const hintFor = async (method: string, path: string): Promise<string | undefined> => {
  const address = OTHER_MAILBOX.exec(path)?.[1]
  if (address === undefined) return undefined
  if ((await tokenKind()) !== "delegated") return undefined
  const scopes = await tokenScopes()
  if (scopes.length === 0) return undefined

  const mailbox = oneLine(safeDecode(address))
  const bare = path.split("?")[0] ?? path
  const sends = method.toUpperCase() === "POST" && SENDS.test(bare)
  const needs = method.toUpperCase() === "GET" ? READ : sends ? SEND : WRITE
  const held = heldScope(scopes, needs)
  if (held === undefined) {
    const doing = sends ? "Sending from" : method.toUpperCase() === "GET" ? "Reading" : "Changing"
    return (
      `The signed-in token has no ${needs.join(" or ")}, which ${doing.toLowerCase()} another person's ` +
      "mailbox needs. A tenant admin adds it to the app's consent, and the server's scope list must include it."
    )
  }
  return sends
    ? `The token holds ${held}, so the permission is in place and Exchange refused the send itself. ` +
        `${mailbox} has to give you Send As or Send on Behalf permission.`
    : `The token holds ${held}, so the permission is in place and Exchange refused the mailbox itself. ` +
        `${mailbox} has to give you access to it, in Outlook or by an admin with Add-MailboxPermission.`
}

/** Adds the cause to a refusal on another person's mailbox, when the token proves it. */
export const explainMailboxRefusal = async <T>(
  method: string,
  path: string,
  result: Either<GraphApiError, T>,
): Promise<Either<GraphApiError, T>> => {
  if (result.isRight()) return result
  const error = result.value as GraphApiError
  if (error.type !== "forbidden") return result
  const hint = await hintFor(method, path)
  return hint === undefined ? result : Left<GraphApiError, T>({ ...error, message: `${error.message} ${hint}` })
}
