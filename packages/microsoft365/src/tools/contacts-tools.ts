import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphContact, ODataResponse } from "../types"
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

export const searchContacts = async (params: { query: string; top?: number }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchContacts(params.query, { $top: params.top ?? 25 })
  return result
    .mapLeft((error) => new UserError(`Failed to search contacts: ${error.message}`))
    .map((response) => formatContactList((response as ODataResponse<never>).value))
}
