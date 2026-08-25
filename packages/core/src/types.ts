// Generic Microsoft Graph infrastructure types shared across Graph MCP servers.
// Domain entity types (messages, events, etc.) live in each server, not here.

export type GraphApiError = {
  readonly type: "network" | "parse" | "api" | "auth" | "throttle" | "not_found" | "forbidden" | "unknown"
  readonly message: string
  readonly status?: number
  readonly graphErrorCode?: string
  /**
   * `error.innerError.code`. Some Graph endpoints return one outer code for several distinct
   * conditions — the transcript endpoints answer both "you lack the scope" and "the tenant disabled
   * this feature" with code "Forbidden" — and only the inner code separates them. Callers that need
   * to react differently, rather than just report, cannot do it from `graphErrorCode` alone.
   */
  readonly graphInnerErrorCode?: string
  readonly retryAfter?: number
}

export type AuthError = {
  readonly type: "config" | "credential" | "token" | "scope"
  readonly message: string
}

export type GraphApiVersion = "v1.0" | "beta"

export type ODataResponse<T> = {
  readonly "@odata.context"?: string
  readonly "@odata.nextLink"?: string
  readonly "@odata.count"?: number
  readonly value: ReadonlyArray<T>
}

export type ODataParams = {
  readonly $select?: ReadonlyArray<string>
  readonly $filter?: string
  readonly $expand?: ReadonlyArray<string>
  readonly $orderby?: string
  readonly $top?: number
  readonly $skip?: number
  readonly $search?: string
  readonly $count?: boolean
}

export type GraphDriveItem = {
  readonly id: string
  readonly name?: string
  readonly size?: number
  readonly lastModifiedDateTime?: string
  readonly webUrl?: string
  readonly createdBy?: { readonly user?: { readonly displayName?: string } }
  readonly lastModifiedBy?: { readonly user?: { readonly displayName?: string } }
  readonly folder?: { readonly childCount?: number }
  readonly file?: { readonly mimeType?: string }
  readonly "@microsoft.graph.downloadUrl"?: string
}
