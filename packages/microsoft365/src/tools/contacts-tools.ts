import { mkdir, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import { isAbsolute, join, resolve } from "node:path"

import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphContact, GraphContactFolder, ODataResponse } from "../types"
import { formatContactDetail, formatContactList } from "../utils/formatters"

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
  return result
    .mapLeft((error) => new UserError(`Failed to create contact: ${error.message}`))
    .map((c) => `Contact created.\n\n${formatContactDetail(c)}`)
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
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

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

  if (Object.keys(updates).length === 0) return Left(new UserError("No fields to update."))

  const result = await client.updateContact(params.contact_id, updates)
  return result
    .mapLeft((error) => new UserError(`Failed to update contact: ${error.message}`))
    .map((c) => `Contact updated.\n\n${formatContactDetail(c)}`)
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
  return result
    .mapLeft((error) => new UserError(`Failed to delete contact: ${error.message}`))
    .map(() => `Contact deleted: ${contact.displayName ?? params.contact_id}\n\n${formatContactDetail(contact)}`)
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

type FolderExport = { readonly id: string | null; readonly path: string; readonly count: number }

export const exportContacts = async (params: { out_dir?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const odataParams = { $select: EXPORT_FIELDS, $top: PAGE_SIZE }

  // /me/contacts is only the default folder. Contacts filed into folders (and folders
  // within folders) are elsewhere, and an export that misses them looks complete.
  type FoundFolder = { readonly folder: GraphContactFolder; readonly path: string }

  // Sequential on purpose: folders are few, and a failure part-way must stop the export
  // rather than leave a partial one that looks complete.
  const walkFolders = async (
    parentPath: string,
    listPath: string,
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
        if (acc.isLeft()) return acc
        const path = `${parentPath}/${folder.displayName ?? folder.id}`
        const children = await walkFolders(path, `/me/contactFolders/${folder.id}/childFolders`)
        return children.map((found) => [...(acc.value as ReadonlyArray<FoundFolder>), { folder, path }, ...found])
      },
      Promise.resolve(Right<UserError, ReadonlyArray<FoundFolder>>([])),
    )
  }

  const folders = await walkFolders("Contacts", "/me/contactFolders")
  if (folders.isLeft()) return Left(folders.value as UserError)

  // Graph's contact-folder listings are not documented as strictly one level deep, so a
  // folder reached twice is exported once rather than trusting the hierarchy.
  const uniqueFolders = (folders.value as ReadonlyArray<FoundFolder>).filter(
    (entry, index, all) => all.findIndex((other) => other.folder.id === entry.folder.id) === index,
  )

  const sources = [
    { id: null, path: "Contacts", listPath: "/me/contacts" },
    ...uniqueFolders.map(({ folder, path }) => ({
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
        summary: [...summary, { id: source.id, path: source.path, count: items.length }],
      })
    },
    Promise.resolve(Right<UserError, Exported>({ contacts: [], summary: [] })),
  )
  if (exported.isLeft()) return Left(exported.value as UserError)
  const { contacts, summary } = exported.value as Exported

  const exportedAt = new Date().toISOString()
  const dir = params.out_dir ? (isAbsolute(params.out_dir) ? params.out_dir : resolve(params.out_dir)) : tmpdir()
  const target = join(dir, `contacts-${exportedAt.replace(/[:.]/g, "-")}.json`)

  try {
    await mkdir(dir, { recursive: true })
    // wx: never overwrite an existing file. 0600: the export is the whole address book.
    await writeFile(target, JSON.stringify({ exportedAt, folders: summary, contacts }, null, 2), {
      flag: "wx",
      mode: 0o600,
    })
  } catch (err) {
    return Left(new UserError(`Failed to write ${target}: ${err instanceof Error ? err.message : String(err)}`))
  }

  const folderLines = summary.map((f) => `  ${f.path}: ${f.count}`).join("\n")
  return Right(`Saved: ${target}\nContacts: ${contacts.length}\nFolders:\n${folderLines}`)
}
