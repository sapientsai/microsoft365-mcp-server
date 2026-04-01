import { UserError } from "fastmcp"
import { type Either, Left, Right } from "functype/either"

const DEFAULT_TTL_MS = 5 * 60 * 1000 // 5 minutes
const CONFIRM_TTL_MS = parseInt(process.env.MS365_CONFIRM_TTL_MS ?? String(DEFAULT_TTL_MS), 10)

type PendingAction = {
  readonly token: string
  readonly toolName: string
  readonly preview: string
  readonly createdAt: number
  readonly execute: () => Promise<string>
}

const pendingActions = new Map<string, PendingAction>()

const generateToken = (): string => {
  const chars = "abcdefghijklmnopqrstuvwxyz0123456789"
  const segments = [4, 4, 4].map(() =>
    Array.from({ length: 4 }, () => chars[Math.floor(Math.random() * chars.length)]).join(""),
  )
  return segments.join("-")
}

const cleanup = () => {
  const now = Date.now()
  for (const [token, action] of pendingActions) {
    if (now - action.createdAt > CONFIRM_TTL_MS) {
      pendingActions.delete(token)
    }
  }
}

export const createPendingAction = (toolName: string, preview: string, executeFn: () => Promise<string>): string => {
  cleanup()
  const token = generateToken()
  pendingActions.set(token, {
    token,
    toolName,
    preview,
    createdAt: Date.now(),
    execute: executeFn,
  })
  return token
}

export const executePendingAction = async (token: string): Promise<Either<UserError, string>> => {
  cleanup()
  const action = pendingActions.get(token)

  if (!action) {
    return Left(
      new UserError(
        `No pending action found for token "${token}". It may have expired (${CONFIRM_TTL_MS / 1000}s TTL).`,
      ),
    )
  }

  pendingActions.delete(token)

  const now = Date.now()
  if (now - action.createdAt > CONFIRM_TTL_MS) {
    return Left(new UserError(`Action "${action.toolName}" expired. Please retry the original operation.`))
  }

  // eslint-disable-next-line functype/prefer-either -- boundary: executeFn may throw UserError from unwrapResult
  try {
    const result = await action.execute()
    return Right(result)
  } catch (error) {
    return Left(
      new UserError(`Failed to execute ${action.toolName}: ${error instanceof Error ? error.message : String(error)}`),
    )
  }
}

export const formatConfirmationPreview = (toolName: string, params: Record<string, unknown>, token: string): string => {
  const paramLines = Object.entries(params)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `- ${k}: ${typeof v === "string" ? v : JSON.stringify(v)}`)
    .join("\n")

  return `Pending confirmation:

Action: ${toolName}
${paramLines}

Token: ${token}
Call confirm_action with token "${token}" to execute this action.`
}

export const isConfirmWritesEnabled = (): boolean => process.env.MS365_CONFIRM_WRITES !== "false"
