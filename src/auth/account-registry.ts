import type { TokenCredential } from "@azure/identity"
import { type Either, Left, Right } from "functype/either"
import { None, type Option, Some } from "functype/option"

import type { AuthConfig, AuthError } from "../types"
import { createCredential, testCredential } from "./auth-modes"
import { GRAPH_DEFAULT_SCOPE } from "./scopes"

export type Account = {
  readonly id: string
  readonly label: string
  readonly config: AuthConfig
  readonly credential: TokenCredential
}

type MutableRegistry = {
  accounts: Map<string, Account>
  defaultId: Option<string>
}

const registry: MutableRegistry = {
  accounts: new Map(),
  defaultId: None(),
}

export const addAccount = async (
  id: string,
  label: string,
  config: AuthConfig,
): Promise<Either<AuthError, Account>> => {
  const credentialResult = createCredential(config)
  if (credentialResult.isLeft()) return Left(credentialResult.value as AuthError)

  const credential = credentialResult.value as TokenCredential
  const testResult = await testCredential(credential)
  if (testResult.isLeft()) return Left(testResult.value as AuthError)

  const account: Account = { id, label, config, credential }
  registry.accounts.set(id, account)

  if (registry.defaultId.isNone()) {
    registry.defaultId = Some(id)
  }

  return Right(account)
}

export const removeAccount = (id: string): boolean => {
  const deleted = registry.accounts.delete(id)
  if (
    deleted &&
    registry.defaultId.fold(
      () => false,
      (d) => d === id,
    )
  ) {
    const first = registry.accounts.keys().next()
    registry.defaultId = first.done ? None() : Some(first.value)
  }
  return deleted
}

export const getAccount = (id: string): Option<Account> => {
  const account = registry.accounts.get(id)
  return account ? Some(account) : None()
}

export const listAccounts = (): ReadonlyArray<{ id: string; label: string; isDefault: boolean }> =>
  [...registry.accounts.values()].map((a) => ({
    id: a.id,
    label: a.label,
    isDefault: registry.defaultId.fold(
      () => false,
      (d) => d === a.id,
    ),
  }))

export const setDefaultAccount = (id: string): Either<AuthError, true> => {
  if (!registry.accounts.has(id)) {
    return Left<AuthError, true>({ type: "config", message: `Account "${id}" not found` })
  }
  registry.defaultId = Some(id)
  return Right(true as const)
}

export const getDefaultAccount = (): Option<Account> =>
  registry.defaultId.flatMap((id) => {
    const account = registry.accounts.get(id)
    return account ? Some(account) : None()
  })

export const getAccountToken = async (accountId?: string): Promise<Either<AuthError, string>> => {
  const account = accountId ? getAccount(accountId) : getDefaultAccount()

  if (account.isNone()) {
    return Left<AuthError, string>({
      type: "config",
      message: accountId ? `Account "${accountId}" not found` : "No accounts registered. Use add_account first.",
    })
  }

  const acct = account.value as Account
  // eslint-disable-next-line functype/prefer-either -- boundary: credential.getToken throws
  try {
    const token = await acct.credential.getToken(GRAPH_DEFAULT_SCOPE)
    if (!token?.token) {
      return Left<AuthError, string>({ type: "token", message: `Failed to acquire token for account "${acct.id}"` })
    }
    return Right(token.token)
  } catch (e) {
    return Left<AuthError, string>({
      type: "token",
      message: `Token acquisition failed for account "${acct.id}": ${String(e)}`,
    })
  }
}

export const getAccountCount = (): number => registry.accounts.size

export const resetRegistry = (): void => {
  registry.accounts.clear()
  registry.defaultId = None()
}
