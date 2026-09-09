// Contacts tool definitions.

import { z } from "zod"

import { createContact, getContact, listContacts, searchContacts } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const contactsTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_contacts",
    description: "List contacts",
    parameters: z.object({
      top: z.number().optional().describe("Number of contacts to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listContacts(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_contact",
    description: "Get detailed contact information",
    parameters: z.object({
      contact_id: z.string().describe("The contact ID"),
    }),
    execute: async (params) => unwrapResult(await getContact(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_contact",
    description: "Create a new contact",
    parameters: z.object({
      given_name: z.string().describe("First name"),
      surname: z.string().optional().describe("Last name"),
      email: z.string().optional().describe("Email address"),
      mobile_phone: z.string().optional().describe("Mobile phone number"),
      company_name: z.string().optional().describe("Company name"),
      job_title: z.string().optional().describe("Job title"),
    }),
    execute: async (params) => unwrapResult(await createContact(params)),
    domain: "contacts",
    readOnly: false,
  },
  {
    name: "search_contacts",
    description: "Search contacts by name or email",
    parameters: z.object({
      query: z.string().describe("Search query"),
      top: z.number().optional().describe("Number of results (default: 25)"),
    }),
    execute: async (params) => unwrapResult(await searchContacts(params)),
    domain: "contacts",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
]
