import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAuthStatus, setAccessToken } from "../auth"
import { getContextToken } from "../auth/token-context"
import { listAccounts, setDefaultAccount } from "../auth/account-registry"
import { formatAuthStatus } from "../utils/formatters"

export const getAuthStatusTool = async (): Promise<Either<UserError, string>> => {
  const result = await getAuthStatus()
  if (result.isLeft()) return Left(new UserError(`Auth error: ${(result.value as { message: string }).message}`))
  return Right(
    formatAuthStatus(
      result.value as { mode: string; authenticated: boolean; scopes: ReadonlyArray<string>; expiresAt?: string },
    ),
  )
}

export const setAccessTokenTool = (params: {
  access_token: string
  expires_on?: string
}): Either<UserError, string> => {
  const expiresOn = params.expires_on ? new Date(params.expires_on) : undefined
  const result = setAccessToken(params.access_token, expiresOn)
  if (result.isLeft())
    return Left(new UserError(`Failed to set token: ${(result.value as { message: string }).message}`))
  return Right("Access token updated successfully.")
}

// eslint-disable-next-line @typescript-eslint/require-await -- FastMCP requires async execute
export const listAccountsTool = async (): Promise<Either<UserError, string>> => {
  const accounts = listAccounts()

  if (accounts.length === 0) {
    const contextToken = getContextToken()
    if (contextToken) {
      return Right("OAuth proxy mode — authenticated user is determined per-request via OAuth. Use `get_me` to see the current user.")
    }
    return Right("No accounts registered.")
  }

  const lines = accounts.map((a) => `- **${a.label}** (${a.id})${a.isDefault ? " [default]" : ""}`)
  return Right(`# Accounts\n\n${lines.join("\n")}`)
}

// eslint-disable-next-line @typescript-eslint/require-await -- FastMCP requires async execute
export const switchAccountTool = async (params: { account_id: string }): Promise<Either<UserError, string>> => {
  const result = setDefaultAccount(params.account_id)
  if (result.isLeft())
    return Left(new UserError(`Failed to switch account: ${(result.value as { message: string }).message}`))
  return Right(`Default account switched to "${params.account_id}".`)
}
