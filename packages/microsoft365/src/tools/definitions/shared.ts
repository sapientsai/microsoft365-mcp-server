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
