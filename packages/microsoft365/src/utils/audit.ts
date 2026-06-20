import type { DirectLogger } from "functype-log"
import { createDirectConsoleLogger } from "functype-log"

const logger: DirectLogger = createDirectConsoleLogger({ prefix: "[ms365-audit]" })

export const auditToolCall = (toolName: string, params: Record<string, unknown>): void => {
  const sanitized = sanitizeParams(params)
  logger.info(`tool:${toolName}`, { tool: toolName, params: sanitized })
}

export const auditToolResult = (toolName: string, success: boolean, durationMs: number): void => {
  const log = success ? logger.info : logger.warn
  log(`result:${toolName}`, { tool: toolName, success, durationMs })
}

export const auditToolError = (toolName: string, error: string): void => {
  logger.error(`error:${toolName}`, { tool: toolName, error })
}

export const auditAuth = (event: string, metadata?: Record<string, unknown>): void => {
  logger.info(`auth:${event}`, metadata)
}

const REDACTED_KEYS = new Set(["access_token", "password", "secret", "content_type"])

const sanitizeParams = (params: Record<string, unknown>): Record<string, unknown> =>
  Object.fromEntries(Object.entries(params).map(([k, v]) => [k, REDACTED_KEYS.has(k) ? "[REDACTED]" : v]))
