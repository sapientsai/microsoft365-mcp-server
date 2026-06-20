import type { z } from "zod"

import type { ToolDomain } from "./tool-registry"

export type ToolDefinition = {
  readonly name: string
  readonly description: string
  readonly parameters: z.ZodType
  readonly execute: (...args: ReadonlyArray<never>) => Promise<string>
  readonly domain: ToolDomain
  readonly readOnly: boolean
  readonly annotations?: {
    readonly readOnlyHint?: boolean
    readonly destructiveHint?: boolean
    readonly openWorldHint?: boolean
  }
}
