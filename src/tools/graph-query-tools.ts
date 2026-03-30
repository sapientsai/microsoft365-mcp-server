import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphApiVersion } from "../types"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const graphQuery = async (params: {
  method: string
  path: string
  body?: string
  version?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const body = params.body ? (JSON.parse(params.body) as Record<string, unknown>) : undefined
  const version = (params.version as GraphApiVersion) ?? undefined

  const result = await client.graphQuery(params.method, params.path, body, version)
  return result
    .mapLeft((error) => new UserError(`Graph query failed: ${error.message}`))
    .map((data) => JSON.stringify(data, null, 2))
}
