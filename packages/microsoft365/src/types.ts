// Generic Graph infrastructure types are owned by @sapientsai/ms-graph-core and
// re-exported here so existing `from "../types"` / `from "./types"` imports keep working.
export type {
  AuthError,
  GraphApiError,
  GraphApiVersion,
  GraphDriveItem,
  ODataParams,
  ODataResponse,
} from "@sapientsai/ms-graph-core"

// Auth mode
export type AuthMode = "interactive" | "certificate" | "client-secret" | "client-token" | "oauth-proxy"

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
  | {
      readonly mode: "oauth-proxy"
      readonly tenantId: string
      readonly clientId: string
      readonly clientSecret: string
      readonly baseUrl: string
    }

// Auth status (returned by get_auth_status tool)
export type AuthStatus = {
  readonly mode: AuthMode
  readonly authenticated: boolean
  readonly scopes: ReadonlyArray<string>
  readonly expiresAt?: string
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

export type GraphMeetingTimeSuggestion = {
  readonly confidence?: number
  readonly order?: number
  readonly organizerAvailability?: string
  readonly suggestionReason?: string
  readonly attendeeAvailability?: ReadonlyArray<{
    readonly availability?: string
    readonly attendee?: { readonly emailAddress?: { readonly address?: string } }
  }>
  readonly meetingTimeSlot?: {
    readonly start?: { readonly dateTime?: string; readonly timeZone?: string }
    readonly end?: { readonly dateTime?: string; readonly timeZone?: string }
  }
}

export type GraphMeetingTimeSuggestionsResult = {
  readonly emptySuggestionsReason?: string
  readonly meetingTimeSuggestions?: ReadonlyArray<GraphMeetingTimeSuggestion>
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

export type GraphSite = {
  readonly id: string
  readonly displayName?: string
  readonly name?: string
  readonly webUrl?: string
  readonly description?: string
  readonly createdDateTime?: string
  readonly lastModifiedDateTime?: string
  readonly siteCollection?: { readonly hostname?: string }
}

export type GraphDrive = {
  readonly id: string
  readonly name?: string
  readonly description?: string
  readonly webUrl?: string
  readonly driveType?: string
  readonly quota?: {
    readonly total?: number
    readonly used?: number
    readonly remaining?: number
  }
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

export type GraphChat = {
  readonly id: string
  readonly topic?: string
  readonly chatType?: string
  readonly createdDateTime?: string
  readonly lastUpdatedDateTime?: string
  readonly members?: ReadonlyArray<{
    readonly displayName?: string
    readonly userId?: string
  }>
}

export type GraphChatMessage = {
  readonly id: string
  readonly body?: { readonly contentType?: string; readonly content?: string }
  readonly from?: { readonly user?: { readonly displayName?: string; readonly id?: string } }
  readonly createdDateTime?: string
  readonly messageType?: string
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
