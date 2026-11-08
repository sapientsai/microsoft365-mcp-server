// Through the index, so a test that mocks "../src/auth" reaches these reads too.
import { getAccessToken, getAuthStatus } from "./index"
import { parseJwtKind, parseJwtTenant, type TokenKind } from "./token-scopes"

// What the signed-in token itself proves, for error messages that name a cause. Each read is made
// only after Graph has refused something, so a working call costs nothing extra. Anything that
// can't be read comes back empty, and the caller then keeps Graph's own words.

/** The token's delegated scopes or app roles; empty when they can't be read. */
export const tokenScopes = async (): Promise<ReadonlyArray<string>> =>
  (await getAuthStatus()).fold(
    () => [] as ReadonlyArray<string>,
    (status) => status.scopes,
  )

/** The tenant the token was issued in, lower-cased; undefined when it can't be read. */
export const tokenTenant = async (): Promise<string | undefined> =>
  (await getAccessToken()).fold(
    () => undefined,
    (token) => parseJwtTenant(token),
  )

/** The first of `wanted` the token holds, compared without case; undefined when it holds none. */
export const heldScope = (scopes: ReadonlyArray<string>, wanted: ReadonlyArray<string>): string | undefined =>
  wanted.find((w) => scopes.some((s) => s.toLowerCase() === w.toLowerCase()))

/** Whether the token acts for a user, as an application, or can't be read. */
export const tokenKind = async (): Promise<TokenKind> =>
  (await getAccessToken()).fold(
    () => "unknown" as TokenKind,
    (token) => parseJwtKind(token),
  )
