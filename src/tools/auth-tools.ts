import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAuthStatus, setAccessToken } from "../auth"
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
