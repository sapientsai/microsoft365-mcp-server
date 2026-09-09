// Chat tool definitions.

import { z } from "zod"

import { listChatMessages, listChats, sendChatMessage } from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const chatsTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_chats",
    description:
      "List your Teams chats (1:1, group, and meeting chats). Note: the self-chat (notes to self) is not listed here — use chat_id '48:notes' to send to it directly.",
    parameters: z.object({
      top: z.number().optional().describe("Number of chats to return (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChats(params)),
    domain: "chats" as const,
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_chat_messages",
    description: "List messages in a Teams chat",
    parameters: z.object({
      chat_id: z.string().describe("Chat ID"),
      top: z.number().optional().describe("Number of messages (default: 25)"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listChatMessages(params)),
    domain: "chats" as const,
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_chat_message",
    description:
      "Send a message in a Teams chat. Use chat_id '48:notes' to send to the user's self-chat (notes to self).",
    parameters: z.object({
      chat_id: z.string().describe("Chat ID. Use '48:notes' for the user's self-chat."),
      content: z.string().describe("Message content"),
    }),
    execute: async (params) => unwrapResult(await sendChatMessage(params)),
    domain: "chats" as const,
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
