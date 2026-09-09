// Teams tool definitions.

import { z } from "zod"

import { listChannelMessages, listChannels, listTeams, sendChannelMessage } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const teamsTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_teams",
    description: "List teams you are a member of",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listTeams(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_channels",
    description: "List channels in a team",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChannels(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_channel_messages",
    description: "List recent messages in a channel",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      channel_id: z.string().describe("Channel ID"),
      top: z.number().optional().describe("Number of messages (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChannelMessages(params)),
    domain: "teams",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_channel_message",
    description: "Send a message to a Teams channel",
    parameters: z.object({
      team_id: z.string().describe("Team ID"),
      channel_id: z.string().describe("Channel ID"),
      content: z.string().describe("Message content"),
    }),
    execute: async (params) => unwrapResult(await sendChannelMessage(params)),
    domain: "teams",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
