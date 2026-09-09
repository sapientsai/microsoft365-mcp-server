// Helpers shared by every per-domain tool-definition module.

import type { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { z } from "zod"

/**
 * FastMCP signals a tool failure by a thrown UserError, so the Either the tool
 * layer returns has to be collapsed at exactly this boundary. Everything below
 * stays in the Either world.
 */
/* eslint-disable functype/prefer-either -- deliberate Either → throw boundary for FastMCP */
export const unwrapResult = <T>(result: Either<UserError, T>): T =>
  result.fold(
    (e) => {
      throw e
    },
    (v) => v,
  )
/* eslint-enable functype/prefer-either */

export const FETCH_ALL_PAGES_PARAM = z.boolean().optional().describe("Fetch all pages of results (max 50 pages)")

// Mail tools accept an optional mailbox so one server can serve a household: the
// signed-in user's own mail by default, a delegated or shared mailbox when named.
// Which addresses are permitted is deployment configuration (MS365_ALLOWED_MAILBOXES),
// not something the caller can widen, so the description points at the error rather
// than listing addresses that would go stale.
export const MAILBOX_PARAM = z
  .string()
  .optional()
  .describe(
    "Email address of another mailbox to act on (delegated or shared), e.g. 'someone@example.com'. " +
      "Omit for your own mailbox. Only addresses this server is configured to allow will be accepted.",
  )
