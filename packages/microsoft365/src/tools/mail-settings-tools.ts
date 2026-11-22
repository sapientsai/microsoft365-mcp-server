import type { Either } from "functype/either"
import { Left, Right } from "functype/either"
import { UserError } from "somamcp"

import type { GraphApiError, GraphMessageRule, GraphOutlookCategory, ODataResponse } from "../types"
import {
  oneLine,
  resolveDestination,
  resolveMessageId,
  type ScopedClient,
  scopeFailure,
  WELL_KNOWN_FOLDERS,
  withScope,
} from "./mail-tools"

const WELL_KNOWN_IDS: ReadonlySet<string> = new Set(WELL_KNOWN_FOLDERS.values())

// Automatic replies, Outlook categories, and inbox rules. Settings live in the mailbox, so these
// need MailboxSettings.ReadWrite; tagging a message with categories needs only Mail.ReadWrite.
//
// Inbox rules here never forward or redirect mail. A forwarding rule sends every matching message
// to someone else with no draft step, so it is left out by design, and graph_query refuses one
// while MS365_REQUIRE_DRAFT is on.

const MAX_TEXT = 300

// Graph answers a token without MailboxSettings with a bare 403 "Access is denied", which doesn't
// say what is missing. Settings are the one thing these calls need beyond the mail scopes.
// For another mailbox a 403 has other causes (mailbox access, or Graph refusing delegated access to
// someone else's settings), so the hint is given only for the caller's own.
const settingsFailure = (action: string, error: GraphApiError, prefix: string): UserError =>
  new UserError(
    error.type === "forbidden" && prefix === "/me"
      ? `Failed to ${action}: ${error.message}. This needs the MailboxSettings.ReadWrite permission: grant it ` +
          "on the app registration, then sign in again."
      : `Failed to ${action}: ${error.message}`,
  )

// ---------- Automatic replies ----------

const dateTimeLine = (label: string, value?: { dateTime: string; timeZone: string }): string =>
  value ? `\n- ${label}: ${value.dateTime} (${value.timeZone})` : ""

export const getAutomaticReplies = async (params: { mailbox?: string }): Promise<Either<UserError, string>> => {
  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const result = await client.getAutomaticReplies(scope.prefix)
  return result
    .mapLeft((error) => settingsFailure("read automatic replies", error, scope.prefix))
    .map(
      (setting) =>
        `# Automatic replies\n\n- Status: ${setting.status ?? "unknown"}` +
        `\n- External audience: ${setting.externalAudience ?? "unknown"}${dateTimeLine(
          "Starts",
          setting.scheduledStartDateTime,
        )}${dateTimeLine(
          "Ends",
          setting.scheduledEndDateTime,
        )}\n- Internal reply: ${oneLine(setting.internalReplyMessage ?? "", 2000) || "(none)"}` +
        `\n- External reply: ${oneLine(setting.externalReplyMessage ?? "", 2000) || "(none)"}`,
    )
}

// A naive date-time with no offset, the form Graph's dateTimeTimeZone takes alongside a time zone.
const NAIVE_DATE_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?$/

export const setAutomaticReplies = async (params: {
  status: "disabled" | "alwaysEnabled" | "scheduled"
  internal_message?: string
  external_message?: string
  external_audience?: "none" | "contactsOnly" | "all"
  start?: string
  end?: string
  time_zone?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  if (params.status === "scheduled" && (params.start === undefined || params.end === undefined)) {
    return Left(new UserError("status scheduled needs both start and end."))
  }
  const badTime = [params.start, params.end].find((value) => value !== undefined && !NAIVE_DATE_TIME.test(value))
  if (badTime !== undefined) {
    return Left(
      new UserError(`"${badTime.slice(0, 40)}" is not a date-time like 2026-12-20T09:00. Put the zone in time_zone.`),
    )
  }
  // Both are naive times in the same zone, so reading them as UTC compares them correctly.
  if (
    params.start !== undefined &&
    params.end !== undefined &&
    Date.parse(`${params.end}Z`) <= Date.parse(`${params.start}Z`)
  ) {
    return Left(new UserError("end must be after start."))
  }

  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const hasTimes = params.start !== undefined || params.end !== undefined
  const zone = await scheduleZone(params.time_zone, hasTimes, client, scope.prefix)
  const timeZone = zone.name
  // Only what the caller passed is sent, so leaving out a message keeps the one already set.
  const setting = {
    status: params.status,
    ...(params.internal_message !== undefined ? { internalReplyMessage: params.internal_message } : {}),
    ...(params.external_message !== undefined ? { externalReplyMessage: params.external_message } : {}),
    ...(params.external_audience !== undefined ? { externalAudience: params.external_audience } : {}),
    ...(params.start !== undefined ? { scheduledStartDateTime: { dateTime: params.start, timeZone } } : {}),
    ...(params.end !== undefined ? { scheduledEndDateTime: { dateTime: params.end, timeZone } } : {}),
  }

  const result = await client.setAutomaticReplies(setting, scope.prefix)
  if (result.isLeft()) {
    const error = result.value as GraphApiError
    // A zone the caller didn't choose may be one Graph won't take for a schedule (a custom zone, or
    // an IANA name outside its list). A 400 can have other causes too, so Graph's own message stays.
    if (zone.source === "mailbox" && error.status === 400) {
      return Left(
        new UserError(
          `Graph refused the change: ${oneLine(error.message, 300)} ` +
            `If the cause is the time zone, the schedule used the mailbox's own, ${JSON.stringify(oneLine(timeZone, 80))}; ` +
            'pass time_zone, such as "Eastern Standard Time" or "UTC".',
        ),
      )
    }
    return Left(settingsFailure("set automatic replies", error, scope.prefix))
  }
  // Graph reports the schedule back in UTC whatever zone it was set in, so the caller is told which
  // zone start and end were read in.
  const zoneNote = hasTimes ? `\n\n${zoneLine(zone)}` : ""
  return getAutomaticReplies({ mailbox: params.mailbox }).then((read) =>
    read.fold(
      () => Right(`Automatic replies set to ${params.status}.${zoneNote}`),
      (text) => Right(`Automatic replies updated.\n\n${text}${zoneNote}`),
    ),
  )
}

type ScheduleZone = { readonly name: string; readonly source: "given" | "mailbox" | "fallback" }

// Left out, the schedule is in the mailbox's own time zone, as Outlook shows it to its owner. A
// window read as UTC would start hours early or late for anyone outside UTC.
const scheduleZone = async (
  given: string | undefined,
  hasTimes: boolean,
  client: ScopedClient["client"],
  prefix: string,
): Promise<ScheduleZone> => {
  if (given !== undefined) return { name: given, source: "given" }
  if (!hasTimes) return { name: "UTC", source: "fallback" }
  return (await client.getMailboxTimeZone(prefix)).fold<ScheduleZone>(
    () => ({ name: "UTC", source: "fallback" }),
    (setting) =>
      setting.value === undefined || setting.value.trim() === ""
        ? { name: "UTC", source: "fallback" }
        : { name: setting.value, source: "mailbox" },
  )
}

const zoneLine = (zone: ScheduleZone): string => {
  const name = JSON.stringify(oneLine(zone.name, 80))
  if (zone.source === "mailbox") return `Start and end were read in the mailbox's time zone, ${name}.`
  if (zone.source === "fallback") {
    return "Start and end were read as UTC: the mailbox's time zone couldn't be read. Pass time_zone to choose one."
  }
  return `Start and end were read in ${name}.`
}

// ---------- Categories ----------

// Outlook's names for the preset colours (desktop client), so a caller can say "green".
const COLOR_NAMES = [
  "red",
  "orange",
  "brown",
  "yellow",
  "green",
  "teal",
  "olive",
  "blue",
  "purple",
  "cranberry",
  "steel",
  "darksteel",
  "gray",
  "darkgray",
  "black",
  "darkred",
  "darkorange",
  "darkbrown",
  "darkyellow",
  "darkgreen",
  "darkteal",
  "darkolive",
  "darkblue",
  "darkpurple",
  "darkcranberry",
] as const

export const CATEGORY_COLORS: ReadonlyArray<string> = ["none", ...COLOR_NAMES]

const presetOf = (color: string | undefined): string | undefined => {
  if (color === undefined) return "none"
  const normalized = color
    .trim()
    .toLowerCase()
    .replace(/[\s_-]/g, "")
  if (normalized === "none") return "none"
  if (/^preset([0-9]|1[0-9]|2[0-4])$/.test(normalized)) return normalized
  const index = COLOR_NAMES.indexOf(normalized as (typeof COLOR_NAMES)[number])
  return index === -1 ? undefined : `preset${index}`
}

const colorName = (preset: string | undefined): string => {
  const match = /^preset(\d+)$/i.exec(preset ?? "")
  return match ? (COLOR_NAMES.at(Number(match[1])) ?? preset ?? "none") : (preset ?? "none")
}

export const listMailCategories = async (params: { mailbox?: string }): Promise<Either<UserError, string>> => {
  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const result = await client.listMasterCategories(scope.prefix)
  return result
    .mapLeft((error) => settingsFailure("list categories", error, scope.prefix))
    .map((response) => {
      const categories = (response as ODataResponse<GraphOutlookCategory>).value
      if (categories.length === 0) return "No categories defined."
      const lines = categories.map(
        (c) => `- ${oneLine(c.displayName ?? "(unnamed)", MAX_TEXT)} (${colorName(c.color)})`,
      )
      return `# Categories (${categories.length})\n\n${lines.join("\n")}`
    })
}

export const createMailCategory = async (params: {
  name: string
  color?: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const name = params.name.trim()
  if (name === "") return Left(new UserError("name is blank."))
  const color = presetOf(params.color)
  if (color === undefined) {
    return Left(new UserError(`Unknown color "${params.color}". Use one of: ${CATEGORY_COLORS.join(", ")}.`))
  }

  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const result = await client.createMasterCategory(name, color, scope.prefix)
  return result
    .mapLeft((error) => settingsFailure("create category", error, scope.prefix))
    .map((c) => `Category created: ${oneLine(c.displayName ?? name, MAX_TEXT)} (${colorName(c.color)})`)
}

const sameName = (a: string, b: string) => a.toLowerCase() === b.toLowerCase()

export const setMessageCategories = async (params: {
  message_id: string
  categories: ReadonlyArray<string>
  mode?: "replace" | "add" | "remove"
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  const requested = params.categories.map((c) => c.trim()).filter((c) => c.length > 0)
  const mode = params.mode ?? "replace"
  if (requested.length === 0 && mode !== "replace") {
    return Left(new UserError(`categories is empty, so mode ${mode} would change nothing.`))
  }

  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const messageId = resolveMessageId(params.message_id, scope)
  if (messageId.isLeft()) return Left(messageId.value as UserError)
  const id = messageId.orThrow()

  const current = async (): Promise<Either<UserError, ReadonlyArray<string>>> => {
    if (mode === "replace") return Right([])
    const read = await client.getMessageCategories(id, scope.prefix)
    return read
      .mapLeft((error) => new UserError(`Failed to read the message's categories: ${error.message}`))
      .map((message) => message.categories ?? [])
  }
  const existing = await current()
  if (existing.isLeft()) return Left(existing.value as UserError)
  const before = existing.orThrow()

  const next =
    mode === "replace"
      ? requested.filter((c, i) => requested.findIndex((other) => sameName(other, c)) === i)
      : mode === "add"
        ? [...before, ...requested.filter((c) => !before.some((b) => sameName(b, c)))]
        : before.filter((b) => !requested.some((c) => sameName(b, c)))

  const result = await client.updateMessage(id, { categories: next }, scope.prefix)
  return result
    .mapLeft((error) => new UserError(`Failed to set categories: ${error.message}`))
    .map((message) => {
      const shown = (message.categories ?? next).map((c) => oneLine(c, MAX_TEXT))
      return `Categories on "${oneLine(message.subject ?? "(No Subject)", MAX_TEXT)}": ${shown.length > 0 ? shown.join(", ") : "(none)"}`
    })
}

// ---------- Inbox rules ----------

const valueText = (value: unknown): string => {
  if (Array.isArray(value)) {
    // Every entry counts when checking where mail goes, so a cut list says how many were left out.
    const items = value.map((item) => oneLine(valueText(item), MAX_TEXT))
    const shown = items.reduce<ReadonlyArray<string>>(
      (kept, item) => (kept.join(", ").length + item.length <= MAX_TEXT || kept.length === 0 ? [...kept, item] : kept),
      [],
    )
    const hidden = items.length - shown.length
    return hidden > 0 ? `${shown.join(", ")} (+${hidden} more)` : shown.join(", ")
  }
  if (typeof value === "object" && value !== null) {
    const address = (value as { emailAddress?: { address?: string } }).emailAddress?.address
    return address ?? JSON.stringify(value)
  }
  return String(value)
}

const partLines = (label: string, part: Record<string, unknown> | undefined): string => {
  const entries = Object.entries(part ?? {}).filter(([, v]) => v !== null && v !== false && v !== undefined)
  if (entries.length === 0) return ""
  return `\n  ${label}: ${entries.map(([k, v]) => `${k}=${Array.isArray(v) ? valueText(v) : oneLine(valueText(v), MAX_TEXT)}`).join("; ")}`
}

// A rule names its target folder by ID. Shown with its name, looked up once per folder; an ID that
// can't be resolved (a deleted folder, a lookup that failed) is shown as is.
type FolderNames = ReadonlyMap<string, string>
const FOLDER_ACTIONS = ["moveToFolder", "copyToFolder"] as const

const folderIdsOf = (rules: ReadonlyArray<GraphMessageRule>): ReadonlyArray<string> => [
  ...new Set(
    rules.flatMap((rule) =>
      FOLDER_ACTIONS.map((key) => (rule.actions as Record<string, unknown> | undefined)?.[key]).filter(
        (id): id is string => typeof id === "string" && id !== "",
      ),
    ),
  ),
]

// The names are a convenience, so the lookups are kept cheap and harmless: one at a time (Outlook
// allows four concurrent requests per mailbox), at most ten, and none after the first failure other
// than "not found", since a 503 or timeout counts against the circuit breaker every tool shares.
const MAX_FOLDER_LOOKUPS = 10

const resolveFolderNames = async (
  client: ScopedClient["client"],
  ids: ReadonlyArray<string>,
  prefix: string,
): Promise<FolderNames> => {
  const step = async (remaining: ReadonlyArray<string>, names: FolderNames): Promise<FolderNames> => {
    if (remaining.length === 0) return names
    const [id, ...rest] = remaining
    const result = await client.getMailFolder(id, prefix)
    if (result.isLeft()) return (result.value as GraphApiError).type === "not_found" ? step(rest, names) : names
    const { displayName } = result.value as { displayName?: string }
    return step(rest, displayName ? new Map([...names, [id, displayName]]) : names)
  }
  return step(ids.slice(0, MAX_FOLDER_LOOKUPS), new Map())
}

const withFolderNames = (actions: Record<string, unknown> | undefined, names: FolderNames) =>
  actions === undefined
    ? undefined
    : Object.fromEntries(
        Object.entries(actions).map(([key, value]) =>
          (FOLDER_ACTIONS as ReadonlyArray<string>).includes(key) && typeof value === "string" && names.has(value)
            ? // Quoted: a folder name is anyone's choice in a shared mailbox, and "; forwardTo=..." in
              // one must not read as another action on the rule.
              [key, JSON.stringify(names.get(value))]
            : [key, value],
        ),
      )

const formatRule = (rule: GraphMessageRule, names: FolderNames = new Map()): string =>
  `- ${oneLine(rule.displayName ?? "(unnamed)", MAX_TEXT)}` +
  ` [${rule.isEnabled === false ? "disabled" : "enabled"}${rule.hasError ? ", has error" : ""}, order ${rule.sequence ?? "?"}]` +
  `\n  ID: ${rule.id}${partLines("When", rule.conditions)}${partLines(
    "Except",
    rule.exceptions,
  )}${partLines("Do", withFolderNames(rule.actions as Record<string, unknown> | undefined, names))}`

export const listMailRules = async (params: { mailbox?: string }): Promise<Either<UserError, string>> => {
  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const result = await client.listMessageRules(scope.prefix)
  if (result.isLeft()) return Left(settingsFailure("list inbox rules", result.value as GraphApiError, scope.prefix))
  const rules = [...(result.value as ReadonlyArray<GraphMessageRule>)].sort(
    (a, b) => (a.sequence ?? 0) - (b.sequence ?? 0),
  )
  if (rules.length === 0) return Right("No inbox rules.")
  const names = await resolveFolderNames(client, folderIdsOf(rules), scope.prefix)
  return Right(`# Inbox rules (${rules.length})\n\n${rules.map((rule) => formatRule(rule, names)).join("\n")}`)
}

export type CreateMailRuleParams = {
  name: string
  from_addresses?: ReadonlyArray<string>
  sender_contains?: ReadonlyArray<string>
  subject_contains?: ReadonlyArray<string>
  body_contains?: ReadonlyArray<string>
  body_or_subject_contains?: ReadonlyArray<string>
  recipient_contains?: ReadonlyArray<string>
  has_attachments?: boolean
  importance?: "low" | "normal" | "high"
  sent_only_to_me?: boolean
  move_to_folder?: string
  copy_to_folder?: string
  mark_as_read?: boolean
  mark_importance?: "low" | "normal" | "high"
  assign_categories?: ReadonlyArray<string>
  delete?: boolean
  stop_processing_rules?: boolean
  sequence?: number
  enabled?: boolean
  mailbox?: string
}

const words = (values: ReadonlyArray<string> | undefined): ReadonlyArray<string> | undefined => {
  const kept = (values ?? []).map((v) => v.trim()).filter((v) => v.length > 0)
  return kept.length > 0 ? kept : undefined
}

const defined = (record: Record<string, unknown>): Record<string, unknown> =>
  Object.fromEntries(Object.entries(record).filter(([, v]) => v !== undefined))

const ruleConditions = (params: CreateMailRuleParams): Record<string, unknown> =>
  defined({
    fromAddresses: words(params.from_addresses)?.map((address) => ({ emailAddress: { address } })),
    senderContains: words(params.sender_contains),
    subjectContains: words(params.subject_contains),
    bodyContains: words(params.body_contains),
    bodyOrSubjectContains: words(params.body_or_subject_contains),
    recipientContains: words(params.recipient_contains),
    // Sent only when true: a false predicate means "not used", so it must not count as a condition.
    hasAttachments: params.has_attachments === true ? true : undefined,
    importance: params.importance,
    sentOnlyToMe: params.sent_only_to_me === true ? true : undefined,
  })

/** A rule's folder action needs a real folder ID: names and well-known names are looked up. */
const ruleFolderId = async (
  scoped: ScopedClient,
  folder: string | undefined,
): Promise<Either<UserError, string | undefined>> => {
  if (folder === undefined) return Right(undefined)
  if (folder.trim() === "") return Left(new UserError("A folder name is blank."))
  const destination = await resolveDestination(scoped.client, folder, scoped.scope)
  if (destination.isLeft()) return Left(destination.value as UserError)
  const target = destination.orThrow()
  // A name matched in the folder list already has its real ID, and an unmatched value is taken as an ID.
  // Only a well-known name such as "archive" needs looking up: a rule won't take one.
  if (target.assumedId || !WELL_KNOWN_IDS.has(target.id)) return Right(target.id)
  const lookup = await scoped.client.getMailFolder(target.id, scoped.scope.prefix)
  return lookup
    .mapLeft((error) => new UserError(`Failed to find folder "${folder}": ${error.message}`))
    .map((found) => found.id)
}

export const createMailRule = async (params: CreateMailRuleParams): Promise<Either<UserError, string>> => {
  const name = params.name.trim()
  if (name === "") return Left(new UserError("name is blank."))
  const conditions = ruleConditions(params)
  if (Object.keys(conditions).length === 0) {
    return Left(
      new UserError(
        "Give at least one condition (from_addresses, sender_contains, subject_contains, body_contains, " +
          "body_or_subject_contains, recipient_contains, has_attachments, importance, sent_only_to_me). " +
          "A rule with none would act on every message.",
      ),
    )
  }
  if (params.delete === true && params.move_to_folder !== undefined) {
    return Left(new UserError("Choose delete or move_to_folder, not both."))
  }

  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const scoped = resolved.orThrow()

  const moveTo = await ruleFolderId(scoped, params.move_to_folder)
  if (moveTo.isLeft()) return Left(moveTo.value as UserError)
  const copyTo = await ruleFolderId(scoped, params.copy_to_folder)
  if (copyTo.isLeft()) return Left(copyTo.value as UserError)

  const actions = defined({
    moveToFolder: moveTo.orThrow(),
    copyToFolder: copyTo.orThrow(),
    markAsRead: params.mark_as_read === true ? true : undefined,
    markImportance: params.mark_importance,
    assignCategories: words(params.assign_categories),
    delete: params.delete === true ? true : undefined,
  })
  if (Object.keys(actions).length === 0) {
    return Left(
      new UserError(
        "Give at least one action: move_to_folder, copy_to_folder, mark_as_read, mark_importance, " +
          "assign_categories or delete.",
      ),
    )
  }

  const rule = {
    displayName: name,
    sequence: params.sequence ?? 1,
    isEnabled: params.enabled ?? true,
    conditions,
    actions: { ...actions, ...(params.stop_processing_rules === true ? { stopProcessingRules: true } : {}) },
  }
  const result = await scoped.client.createMessageRule(rule, scoped.scope.prefix)
  if (result.isLeft()) {
    return Left(settingsFailure("create inbox rule", result.value as GraphApiError, scoped.scope.prefix))
  }
  const created = result.value as GraphMessageRule
  const names = await resolveFolderNames(scoped.client, folderIdsOf([created]), scoped.scope.prefix)
  return Right(`Inbox rule created.\n\n${formatRule(created, names)}`)
}

export const deleteMailRule = async (params: {
  rule_id: string
  mailbox?: string
}): Promise<Either<UserError, string>> => {
  // Encoding keeps a slash inside one segment, but "." and ".." would still be resolved as a path.
  if (params.rule_id.trim() === "" || /^\.+$/.test(params.rule_id.trim())) {
    return Left(new UserError("rule_id is not a rule ID. Use the ID from list_mail_rules."))
  }
  const resolved = await withScope(params.mailbox)
  if (resolved.isLeft()) return scopeFailure(resolved)
  const { scope, client } = resolved.orThrow()

  const result = await client.deleteMessageRule(params.rule_id, scope.prefix)
  return result
    .mapLeft((error) => settingsFailure("delete inbox rule", error, scope.prefix))
    .map(() => `Inbox rule ${oneLine(params.rule_id, MAX_TEXT)} deleted.`)
}
