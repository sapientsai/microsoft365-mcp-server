// Users and groups tool definitions.

import { z } from "zod"

import { getGroup, getMe, getUser, listGroupMembers, listGroups, listUsers } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const usersGroupsTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "get_me",
    description: "Get the authenticated user's profile",
    parameters: z.object({}),
    execute: async () => unwrapResult(await getMe()),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_users",
    description: "List users in the organization",
    parameters: z.object({
      top: z.number().optional().describe("Number of users (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listUsers(params)),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_user",
    description: "Get a specific user's profile",
    parameters: z.object({
      user_id: z.string().describe("User ID or UPN"),
    }),
    execute: async (params) => unwrapResult(await getUser(params)),
    domain: "users",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_groups",
    description: "List groups in the organization",
    parameters: z.object({
      top: z.number().optional().describe("Number of groups (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listGroups(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_group",
    description: "Get detailed group information",
    parameters: z.object({
      group_id: z.string().describe("Group ID"),
    }),
    execute: async (params) => unwrapResult(await getGroup(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_group_members",
    description: "List members of a group",
    parameters: z.object({
      group_id: z.string().describe("Group ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listGroupMembers(params)),
    domain: "groups",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
]
