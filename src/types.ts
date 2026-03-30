// Graph API error types
export type GraphApiError = {
  readonly type: "network" | "parse" | "api" | "auth" | "throttle" | "not_found" | "forbidden" | "unknown"
  readonly message: string
  readonly status?: number
  readonly graphErrorCode?: string
  readonly retryAfter?: number
}

export type AuthError = {
  readonly type: "config" | "credential" | "token" | "scope"
  readonly message: string
}

// Graph API version
export type GraphApiVersion = "v1.0" | "beta"

// Auth mode
export type AuthMode = "interactive" | "certificate" | "client-secret" | "client-token"

// Auth config discriminated union
export type AuthConfig =
  | {
      readonly mode: "interactive"
      readonly tenantId: string
      readonly clientId: string
      readonly redirectUri?: string
    }
  | {
      readonly mode: "certificate"
      readonly tenantId: string
      readonly clientId: string
      readonly certPath: string
      readonly certPassword?: string
    }
  | {
      readonly mode: "client-secret"
      readonly tenantId: string
      readonly clientId: string
      readonly clientSecret: string
    }
  | { readonly mode: "client-token"; readonly accessToken?: string; readonly expiresOn?: Date }

// Auth status (returned by get_auth_status tool)
export type AuthStatus = {
  readonly mode: AuthMode
  readonly authenticated: boolean
  readonly scopes: ReadonlyArray<string>
  readonly expiresAt?: string
}

// OData response wrapper
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

// Microsoft Graph entity types

export type GraphUser = {
  readonly id: string
  readonly displayName?: string
  readonly mail?: string
  readonly userPrincipalName?: string
  readonly jobTitle?: string
  readonly department?: string
  readonly officeLocation?: string
  readonly mobilePhone?: string
  readonly businessPhones?: ReadonlyArray<string>
}

export type GraphMessage = {
  readonly id: string
  readonly subject?: string
  readonly from?: { readonly emailAddress: { readonly name?: string; readonly address?: string } }
  readonly toRecipients?: ReadonlyArray<{
    readonly emailAddress: { readonly name?: string; readonly address?: string }
  }>
  readonly receivedDateTime?: string
  readonly isRead?: boolean
  readonly hasAttachments?: boolean
  readonly bodyPreview?: string
  readonly body?: { readonly contentType?: string; readonly content?: string }
  readonly importance?: string
}

export type GraphEvent = {
  readonly id: string
  readonly subject?: string
  readonly start?: { readonly dateTime?: string; readonly timeZone?: string }
  readonly end?: { readonly dateTime?: string; readonly timeZone?: string }
  readonly location?: { readonly displayName?: string }
  readonly organizer?: { readonly emailAddress: { readonly name?: string; readonly address?: string } }
  readonly attendees?: ReadonlyArray<{
    readonly emailAddress: { readonly name?: string; readonly address?: string }
    readonly type?: string
    readonly status?: { readonly response?: string }
  }>
  readonly isAllDay?: boolean
  readonly isCancelled?: boolean
  readonly body?: { readonly contentType?: string; readonly content?: string }
  readonly onlineMeeting?: { readonly joinUrl?: string }
  readonly recurrence?: unknown
}

export type GraphContact = {
  readonly id: string
  readonly displayName?: string
  readonly givenName?: string
  readonly surname?: string
  readonly emailAddresses?: ReadonlyArray<{ readonly name?: string; readonly address?: string }>
  readonly businessPhones?: ReadonlyArray<string>
  readonly mobilePhone?: string
  readonly companyName?: string
  readonly jobTitle?: string
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

export type GraphTeam = {
  readonly id: string
  readonly displayName?: string
  readonly description?: string
}

export type GraphChannel = {
  readonly id: string
  readonly displayName?: string
  readonly description?: string
  readonly membershipType?: string
}

export type GraphChannelMessage = {
  readonly id: string
  readonly body?: { readonly contentType?: string; readonly content?: string }
  readonly from?: { readonly user?: { readonly displayName?: string } }
  readonly createdDateTime?: string
}

export type GraphGroup = {
  readonly id: string
  readonly displayName?: string
  readonly description?: string
  readonly mail?: string
  readonly groupTypes?: ReadonlyArray<string>
  readonly membershipRule?: string
}

export type GraphPlan = {
  readonly id: string
  readonly title?: string
  readonly owner?: string
  readonly createdDateTime?: string
}

export type GraphPlannerTask = {
  readonly id: string
  readonly title?: string
  readonly planId?: string
  readonly bucketId?: string
  readonly percentComplete?: number
  readonly priority?: number
  readonly dueDateTime?: string
  readonly createdDateTime?: string
  readonly assignments?: Record<string, unknown>
}

export type GraphNotebook = {
  readonly id: string
  readonly displayName?: string
  readonly createdDateTime?: string
  readonly lastModifiedDateTime?: string
  readonly isDefault?: boolean
}

export type GraphSection = {
  readonly id: string
  readonly displayName?: string
  readonly createdDateTime?: string
  readonly lastModifiedDateTime?: string
}

export type GraphPage = {
  readonly id: string
  readonly title?: string
  readonly createdDateTime?: string
  readonly lastModifiedDateTime?: string
  readonly contentUrl?: string
}

export type GraphTodoList = {
  readonly id: string
  readonly displayName?: string
  readonly isOwner?: boolean
  readonly isShared?: boolean
  readonly wellknownListName?: string
}

export type GraphTodoTask = {
  readonly id: string
  readonly title?: string
  readonly status?: string
  readonly importance?: string
  readonly isReminderOn?: boolean
  readonly body?: { readonly contentType?: string; readonly content?: string }
  readonly dueDateTime?: { readonly dateTime?: string; readonly timeZone?: string }
  readonly completedDateTime?: { readonly dateTime?: string; readonly timeZone?: string }
  readonly createdDateTime?: string
  readonly lastModifiedDateTime?: string
}
