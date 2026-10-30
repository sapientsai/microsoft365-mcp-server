import { tmpdir } from "node:os"
import { isAbsolute, join, resolve } from "node:path"

import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getAuthStatus } from "../auth"
import { getGraphClient } from "../client/graph-client"
import type { GraphApiError, GraphContact, GraphContactFolder, ODataResponse } from "../types"
import { formatContactDetail, formatContactList } from "../utils/formatters"
import { writeNewFile } from "../utils/local-file"

// Graph answers a contact write without Contacts.ReadWrite with a bare "Access is denied", which
// reads like a bad sign-in. When the token's own scopes show that is the cause, say so, as driveError
// in files-tools.ts does. A 403 from a token that holds the scope, or whose scopes can't be read, keeps
// Graph's message, since the cause is then something else.
const writeFailure = async (action: string, error: GraphApiError): Promise<UserError> => {
  const plain = new UserError(`Failed to ${action} contact: ${error.message}`)
  if (error.type !== "forbidden") return plain
  const scopes = (await getAuthStatus()).fold(
    () => [] as ReadonlyArray<string>,
    (status) => status.scopes,
  )
  const lacksWrite = scopes.length > 0 && !scopes.some((scope) => scope.toLowerCase() === "contacts.readwrite")
  if (!lacksWrite) return plain
  return new UserError(
    `Failed to ${action} contact: ${error.message}. The signed-in token has no Contacts.ReadWrite permission, ` +
      "so contacts can be read but not changed. Add Contacts.ReadWrite to the app registration and grant consent.",
  )
}

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

export const listContacts = async (params: {
  top?: number
  filter?: string
  fetch_all_pages?: boolean
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.fetch_all_pages) {
    const result = await client.requestPaginated<GraphContact>("/me/contacts", {
      odataParams: { $filter: params.filter },
    })
    return result
      .mapLeft((error) => new UserError(`Failed to list contacts: ${error.message}`))
      .map((items) => formatContactList(items))
  }

  const result = await client.listContacts({ $top: params.top ?? 25, $filter: params.filter })
  return result
    .mapLeft((error) => new UserError(`Failed to list contacts: ${error.message}`))
    .map((response) => formatContactList((response as ODataResponse<never>).value))
}

export const getContact = async (params: { contact_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getContact(params.contact_id)
  return result.mapLeft((error) => new UserError(`Failed to get contact: ${error.message}`)).map(formatContactDetail)
}

export const createContact = async (params: {
  given_name: string
  surname?: string
  email?: string
  mobile_phone?: string
  company_name?: string
  job_title?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const contact: Record<string, unknown> = { givenName: params.given_name }
  if (params.surname) contact.surname = params.surname
  if (params.email) contact.emailAddresses = [{ address: params.email }]
  if (params.mobile_phone) contact.mobilePhone = params.mobile_phone
  if (params.company_name) contact.companyName = params.company_name
  if (params.job_title) contact.jobTitle = params.job_title

  const result = await client.createContact(contact)
  if (result.isLeft()) return Left(await writeFailure("create", result.value as GraphApiError))
  return Right(`Contact created.\n\n${formatContactDetail(result.value as GraphContact)}`)
}

const BIRTHDAY_PATTERN = /^(\d{4})-(\d{2})-(\d{2})(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2}))?$/

// Accepts a timestamp as well as a bare date so a birthday read off another contact can be
// copied verbatim: re-deriving the date from a UTC timestamp can shift it by a day.
// The component check catches dates Date would silently roll over, like 2023-02-30.
const isValidBirthday = (value: string): boolean => {
  const match = BIRTHDAY_PATTERN.exec(value)
  if (!match || Number.isNaN(Date.parse(value))) return false
  const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])]
  const date = new Date(Date.UTC(year, month - 1, day))
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
}

// Graph replaces a multi-valued property wholesale on PATCH — there is no append — so
// list parameters are complete replacements, and an empty list clears the property.
export const updateContact = async (params: {
  contact_id: string
  given_name?: string
  surname?: string
  nickname?: string
  company_name?: string
  job_title?: string
  mobile_phone?: string
  email_addresses?: ReadonlyArray<string>
  business_phones?: ReadonlyArray<string>
  home_phones?: ReadonlyArray<string>
  categories?: ReadonlyArray<string>
  personal_notes?: string
  birthday?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.birthday !== undefined && !isValidBirthday(params.birthday)) {
    return Left(
      new UserError(
        `Invalid birthday "${params.birthday}". Use a date (YYYY-MM-DD) or an ISO 8601 timestamp with a UTC offset.`,
      ),
    )
  }

  const updates: Record<string, unknown> = {}
  if (params.given_name !== undefined) updates.givenName = params.given_name
  if (params.surname !== undefined) updates.surname = params.surname
  if (params.nickname !== undefined) updates.nickName = params.nickname
  if (params.company_name !== undefined) updates.companyName = params.company_name
  if (params.job_title !== undefined) updates.jobTitle = params.job_title
  if (params.mobile_phone !== undefined) updates.mobilePhone = params.mobile_phone
  if (params.email_addresses !== undefined) {
    updates.emailAddresses = params.email_addresses.map((address) => ({ address }))
  }
  if (params.business_phones !== undefined) updates.businessPhones = params.business_phones
  if (params.home_phones !== undefined) updates.homePhones = params.home_phones
  if (params.categories !== undefined) updates.categories = params.categories
  if (params.personal_notes !== undefined) updates.personalNotes = params.personal_notes
  if (params.birthday !== undefined) updates.birthday = params.birthday

  if (Object.keys(updates).length === 0) return Left(new UserError("No fields to update."))

  const result = await client.updateContact(params.contact_id, updates)
  if (result.isLeft()) return Left(await writeFailure("update", result.value as GraphApiError))
  return Right(`Contact updated.\n\n${formatContactDetail(result.value as GraphContact)}`)
}

export const deleteContact = async (params: { contact_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  // Read first so the caller is told what was removed, not just an ID that no longer resolves.
  const existing = await client.getContact(params.contact_id)
  if (existing.isLeft()) {
    return Left(new UserError(`Failed to get contact: ${(existing.value as { message: string }).message}`))
  }
  const contact = existing.value as GraphContact

  const result = await client.deleteContact(params.contact_id)
  if (result.isLeft()) return Left(await writeFailure("delete", result.value as GraphApiError))
  return Right(`Contact deleted: ${contact.displayName ?? params.contact_id}\n\n${formatContactDetail(contact)}`)
}

export const searchContacts = async (params: { query: string; top?: number }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchContacts(params.query, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to search contacts: ${error.message}`))
    .map((response) => formatContactList((response as ODataResponse<never>).value))
}

// Everything a contact carries that a caller might want to reconcile against another record.
// /me/contacts returns a trimmed default set, so the fields have to be named.
const EXPORT_FIELDS = [
  "id",
  "parentFolderId",
  "displayName",
  "givenName",
  "middleName",
  "surname",
  "nickName",
  "emailAddresses",
  "mobilePhone",
  "businessPhones",
  "homePhones",
  "imAddresses",
  "companyName",
  "jobTitle",
  "department",
  "categories",
  "birthday",
  "spouseName",
  "children",
  "homeAddress",
  "businessAddress",
  "otherAddress",
  "personalNotes",
  "createdDateTime",
  "lastModifiedDateTime",
]

// 999 is Graph's page ceiling for contacts. It matters because requestPaginated stops at
// 50 pages without saying so: at the default page size that is 500 contacts, silently
// truncated; at 999 it is ~50,000.
const PAGE_SIZE = 999
// Mirrors MAX_PAGES in core's fetchAllPages, which returns what it has when it hits the cap
// with no sign it stopped early. A folder that fills every page is reported as possibly
// truncated; it may also hold exactly that many contacts, which can't be told apart here.
const MAX_PAGES = 50
const PAGINATION_CAP = PAGE_SIZE * MAX_PAGES

type FolderExport = {
  readonly id: string | null
  readonly path: string
  readonly count: number
  readonly truncated?: true
}

export const exportContacts = async (params: { out_dir?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const odataParams = { $select: EXPORT_FIELDS, $top: PAGE_SIZE }

  // /me/contacts is only the default folder. Contacts filed into folders (and folders
  // within folders) are elsewhere, and an export that misses them looks complete.
  type FoundFolder = { readonly folder: GraphContactFolder; readonly path: string }

  // Sequential on purpose: folders are few, and a failure part-way must stop the export
  // rather than leave a partial one that looks complete.
  // Graph's contact-folder listings are not documented as strictly one level deep, so a
  // folder already reached is skipped rather than trusting the hierarchy: a folder listed
  // twice is exported once, and a folder that lists an ancestor cannot recurse forever.
  const walkFolders = async (
    parentPath: string,
    listPath: string,
    visited: Set<string>,
  ): Promise<Either<UserError, ReadonlyArray<FoundFolder>>> => {
    const listed = await client.requestPaginated<GraphContactFolder>(listPath, { odataParams: { $top: 100 } })
    if (listed.isLeft()) {
      return Left(new UserError(`Failed to list contact folders: ${(listed.value as { message: string }).message}`))
    }
    return (listed.value as ReadonlyArray<GraphContactFolder>).reduce<
      Promise<Either<UserError, ReadonlyArray<FoundFolder>>>
    >(
      async (accPromise, folder) => {
        const acc = await accPromise
        if (acc.isLeft() || visited.has(folder.id)) return acc
        visited.add(folder.id)
        const path = `${parentPath}/${folder.displayName ?? folder.id}`
        const children = await walkFolders(path, `/me/contactFolders/${folder.id}/childFolders`, visited)
        return children.map((found) => [...(acc.value as ReadonlyArray<FoundFolder>), { folder, path }, ...found])
      },
      Promise.resolve(Right<UserError, ReadonlyArray<FoundFolder>>([])),
    )
  }

  const folders = await walkFolders("Contacts", "/me/contactFolders", new Set())
  if (folders.isLeft()) return Left(folders.value as UserError)

  const sources = [
    { id: null, path: "Contacts", listPath: "/me/contacts" },
    ...(folders.value as ReadonlyArray<FoundFolder>).map(({ folder, path }) => ({
      id: folder.id,
      path,
      listPath: `/me/contactFolders/${folder.id}/contacts`,
    })),
  ]

  type Exported = {
    readonly contacts: ReadonlyArray<Record<string, unknown>>
    readonly summary: ReadonlyArray<FolderExport>
  }

  const exported = await sources.reduce<Promise<Either<UserError, Exported>>>(
    async (accPromise, source) => {
      const acc = await accPromise
      if (acc.isLeft()) return acc
      const result = await client.requestPaginated<Record<string, unknown>>(source.listPath, { odataParams })
      if (result.isLeft()) {
        return Left(
          new UserError(`Failed to list contacts in ${source.path}: ${(result.value as { message: string }).message}`),
        )
      }
      const items = result.value as ReadonlyArray<Record<string, unknown>>
      const { contacts, summary } = acc.value as Exported
      return Right({
        contacts: [...contacts, ...items.map((c) => ({ ...c, folderPath: source.path }))],
        summary: [
          ...summary,
          {
            id: source.id,
            path: source.path,
            count: items.length,
            ...(items.length >= PAGINATION_CAP ? { truncated: true as const } : {}),
          },
        ],
      })
    },
    Promise.resolve(Right<UserError, Exported>({ contacts: [], summary: [] })),
  )
  if (exported.isLeft()) return Left(exported.value as UserError)
  const { contacts, summary } = exported.value as Exported

  const exportedAt = new Date().toISOString()
  const dir = params.out_dir ? (isAbsolute(params.out_dir) ? params.out_dir : resolve(params.out_dir)) : tmpdir()
  const name = `contacts-${exportedAt.replace(/[:.]/g, "-")}.json`

  // 0600 matters here: the export is the whole address book.
  const written = await writeNewFile(dir, name, JSON.stringify({ exportedAt, folders: summary, contacts }, null, 2))
  if (written.isLeft()) return Left(new UserError(`Failed to write ${join(dir, name)}: ${written.value.message}`))
  const target = written.value

  const folderLines = summary
    .map((f) => `  ${f.path}: ${f.count}${f.truncated ? " (may be incomplete: hit the paging limit)" : ""}`)
    .join("\n")
  const truncatedPaths = summary.filter((f) => f.truncated).map((f) => f.path)
  const warning =
    truncatedPaths.length > 0
      ? `\n\nWARNING: stopped at the ${PAGINATION_CAP}-contact paging limit in ${truncatedPaths.join(", ")}. ` +
        "The export may be missing contacts from those folders."
      : ""
  return Right(`Saved: ${target}\nContacts: ${contacts.length}\nFolders:\n${folderLines}${warning}`)
}
