// Mail tool definitions.

import { z } from "zod"

import {
  batchMoveMessages,
  createDraft,
  createForwardDraft,
  createReplyAllDraft,
  createReplyDraft,
  getMessage,
  listAttachments,
  listMailFolders,
  listMessages,
  moveMessage,
  saveAttachment,
  scanMessages,
  searchMessages,
  sendDraft,
  sendForward,
  sendMessage,
  sendReply,
  sendReplyAll,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const mailTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_messages",
    description: "List email messages from your inbox",
    parameters: z.object({
      top: z.number().optional().describe("Number of messages to return (default: 25)"),
      filter: z.string().optional().describe("OData filter expression"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "scan_messages",
    description:
      "Scan message headers compactly for triage. Returns pipe-delimited rows (ref|received|from|subject|flags) instead of full markdown — roughly a third the tokens of list_messages — so thousands of messages can be surveyed to decide which few are worth opening. Scope with folder (e.g. 'archive'), narrow with filter or search. Pass a returned ref to get_message in place of the message ID.\n\nCOVERAGE — use FILTER, not SEARCH, whenever completeness matters. A filter pages properly with skip, so you can walk a folder to the end and know you did. A SEARCH cannot be paged at all (Graph ignores skip on search and silently returns page one again, so passing both is rejected) — narrowing a search by date only *shrinks* each window; it never tells you whether the window you got back was whole.\n\nThe trap: a search that returns fewer rows than `top` looks complete and usually is not. Only a filter scan that returns fewer rows than `top` is genuinely exhausted. So for 'have I seen every message with an attachment', use filter: \"hasAttachments eq true and receivedDateTime ge 2024-01-01T00:00:00Z\" and page with skip until a page comes back short. Reserve search for finding known things by keyword.\n\nA page is capped at 999 rows and any truncated result says INCOMPLETE — treat that as 'you have not seen everything', not as an answer.\n\nA header is not a message: a routine-looking subject can carry a substantial attachment. Open what the subject cannot rule out.",
    parameters: z.object({
      folder: z
        .string()
        .optional()
        .describe(
          "Folder to scan: well-known name (archive, inbox, sentitems), display name, or ID. Default: all mail",
        ),
      filter: z
        .string()
        .optional()
        .describe(
          'OData filter, e.g. "receivedDateTime ge 2026-01-01T00:00:00Z" or "hasAttachments eq true". ' +
            "Pageable with skip — use this, not search, when coverage matters.",
        ),
      search: z
        .string()
        .optional()
        .describe(
          "Full-text search term, for finding known things. NOT pageable — cannot be combined with skip or " +
            "date ordering, and a short result does not mean you have seen everything.",
        ),
      top: z.number().optional().describe("Rows per page (default 100, max 999)"),
      skip: z.number().optional().describe("Rows to skip, for paging through large folders"),
    }),
    execute: async (params) => unwrapResult(await scanMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_message",
    description:
      "Get a specific email message with full body content. Pass body_format:'text' for marketing or newsletter mail — Graph converts server-side, avoiding tens of thousands of characters of HTML and CSS.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref returned by scan_messages"),
      body_format: z
        .enum(["text", "html"])
        .optional()
        .describe("Body format to request. 'text' strips HTML/CSS server-side. Default: the message's own format"),
    }),
    execute: async (params) => unwrapResult(await getMessage(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_mail_folders",
    description: "List mail folders with item and unread counts, for resolving move destinations",
    parameters: z.object({
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listMailFolders(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "move_message",
    description:
      "Move a message to another mail folder. Destination accepts a well-known name (archive, deleteditems, inbox, junkemail), a folder display name, or a folder ID. Moving to deleteditems is recoverable; use list_mail_folders to see what exists.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to move"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
    }),
    execute: async (params) => unwrapResult(await moveMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "batch_move_messages",
    description:
      "Move several messages to the same folder in one call. Resolves the destination once and returns a single summary instead of one result per message. Reports any failures individually.",
    parameters: z.object({
      message_ids: z.array(z.string()).describe("Message IDs to move (max 50)"),
      destination: z
        .string()
        .describe("Destination folder: well-known name (e.g. archive), display name, or folder ID"),
    }),
    execute: async (params) => unwrapResult(await batchMoveMessages(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "list_attachments",
    description:
      "List a message's attachments with name, content type and size. Returns a read_document path per " +
      "attachment for extracting its text (PDF, Office, etc). Cloud links (reference attachments to " +
      "OneDrive/SharePoint/Dropbox) are listed with their URL instead — the mailbox holds no bytes for those, " +
      "so they must be opened rather than fetched.",
    parameters: z.object({
      message_id: z.string().describe("The message ID whose attachments to list"),
    }),
    execute: async (params) => unwrapResult(await listAttachments(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "save_attachment",
    description:
      "Save a mail attachment to a local file and return its path. Use for anything read_document cannot " +
      "extract text from — scanned PDFs, photographed documents, images — and for handing a file to another " +
      "tool. Read the saved file directly: PDFs and images are viewable without text extraction. Omit " +
      "attachment_id when the message has exactly one attachment; otherwise get IDs from list_attachments. " +
      "Cloud links (reference attachments) cannot be saved — the tool reports their URL so they can be opened.",
    parameters: z.object({
      message_id: z.string().describe("The message ID, or a short ref returned by scan_messages"),
      attachment_id: z
        .string()
        .optional()
        .describe("Which attachment to save. Optional when the message has exactly one."),
      out_dir: z.string().optional().describe("Directory to write into (default: the system temp directory)"),
    }),
    execute: async (params) => unwrapResult(await saveAttachment(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "send_message",
    description: "Send a new email message",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
    }),
    execute: async (params) => unwrapResult(await sendMessage(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_reply",
    description:
      "Reply to the sender of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await sendReply(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_reply_all",
    description:
      "Reply to all recipients of an email and send immediately. Threads into the conversation and quotes the original. Use create_reply_all_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await sendReplyAll(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "send_forward",
    description:
      "Forward an email to new recipients and send immediately. Quotes the original. Use create_forward_draft to review before sending.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
    }),
    execute: async (params) => unwrapResult(await sendForward(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
  {
    name: "create_reply_draft",
    description:
      "Create a reply draft (to the sender) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await createReplyDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_reply_all_draft",
    description:
      "Create a reply-all draft (to all recipients) in the Drafts folder. Threads into the conversation with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to reply to"),
      comment: z.string().describe("Reply content (added above the quoted original)"),
    }),
    execute: async (params) => unwrapResult(await createReplyAllDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "create_forward_draft",
    description:
      "Create a forward draft in the Drafts folder with the original quoted underneath. Review, then send with send_draft.",
    parameters: z.object({
      message_id: z.string().describe("The message ID to forward"),
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      comment: z.string().optional().describe("Optional note added above the quoted original"),
    }),
    execute: async (params) => unwrapResult(await createForwardDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "search_messages",
    description: "Search email messages",
    parameters: z.object({
      query: z.string().describe("Search query string"),
      top: z.number().optional().describe("Number of results to return (default: 25)"),
    }),
    execute: async (params) => unwrapResult(await searchMessages(params)),
    domain: "mail",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_draft",
    description: "Create a new email draft in the Drafts folder",
    parameters: z.object({
      to: z.string().describe("Recipient email address(es), comma-separated for multiple"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body content"),
      content_type: z.string().optional().describe("Body content type: Text or HTML (default: Text)"),
      cc: z.string().optional().describe("CC recipients (comma-separated email addresses)"),
      bcc: z.string().optional().describe("BCC recipients (comma-separated email addresses)"),
    }),
    execute: async (params) => unwrapResult(await createDraft(params)),
    domain: "mail",
    readOnly: false,
  },
  {
    name: "send_draft",
    description: "Send an existing email draft",
    parameters: z.object({
      message_id: z.string().describe("The draft message ID to send"),
    }),
    execute: async (params) => unwrapResult(await sendDraft(params)),
    domain: "mail",
    readOnly: false,
    annotations: { destructiveHint: true },
  },
]
