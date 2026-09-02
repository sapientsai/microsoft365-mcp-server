// Files and SharePoint tool definitions.

import { z } from "zod"

import {
  createFolder,
  downloadFile,
  getDriveItem,
  getSite,
  getUploadConfig,
  listDriveItems,
  listSiteDrives,
  listSiteItems,
  listSites,
  readDocument,
  searchFiles,
  searchSiteFiles,
  uploadFile,
  uploadFileFromPath,
} from ".."
import type { ToolDefinition } from "../tool-definitions"
import { FETCH_ALL_PAGES_PARAM, unwrapResult } from "./shared"

export const filesTools: ReadonlyArray<ToolDefinition> = [
  {
    name: "list_drive_items",
    description: "List files and folders in OneDrive",
    parameters: z.object({
      folder_id: z.string().optional().describe("Folder ID (omit for root)"),
      folder_path: z.string().optional().describe("Folder path (e.g., 'Documents' or 'Documents/Subfolder')"),
      fetch_all_pages: FETCH_ALL_PAGES_PARAM,
    }),
    execute: async (params) => unwrapResult(await listDriveItems(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_drive_item",
    description: "Get file or folder metadata",
    parameters: z.object({
      item_id: z.string().describe("Drive item ID"),
    }),
    execute: async (params) => unwrapResult(await getDriveItem(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "search_files",
    description: "Search OneDrive/SharePoint files",
    parameters: z.object({
      query: z.string().describe("Search query"),
    }),
    execute: async (params) => unwrapResult(await searchFiles(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "download_file",
    description:
      "Get a file's metadata and download URL. Returns content inline for text files under 100KB. For a " +
      "SharePoint file, pass drive_id as well — without it the item is looked up in your own OneDrive and " +
      "will not be found. For readable text from PDF/DOCX/XLSX, use read_document instead; come here when " +
      "extraction fails (scanned PDFs, unsupported types, files over the extraction cap) or you need raw bytes.",
    parameters: z.object({
      item_id: z.string().describe("Drive item ID"),
      drive_id: z
        .string()
        .optional()
        .describe("Drive ID holding the item. Required for SharePoint; omit for your own OneDrive."),
    }),
    execute: async (params) => unwrapResult(await downloadFile(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "read_document",
    description:
      "Download a file from SharePoint or OneDrive and return its readable text content. Supports DOCX, PDF, XLSX, " +
      "and text-based files. Use instead of download_file when you need document contents. Pair with " +
      "search_site_files (SharePoint) or search_files (OneDrive) to get IDs, then pass " +
      "/drives/{driveId}/items/{itemId}/content or /me/drive/items/{id}/content. Text extraction only, no OCR: " +
      "scanned PDFs return no text — for those, and for images, use save_attachment and read the file.",
    parameters: z.object({
      path: z.string().describe("Graph path to the file content endpoint, ending in /content"),
      api_version: z.enum(["v1.0", "beta"]).optional().describe("Graph API version"),
      format: z.string().optional().describe("Optional conversion format (e.g. 'pdf'), for supported types only"),
      max_chars: z
        .number()
        .int()
        .min(1000)
        .max(200000)
        .optional()
        .describe("Max characters to return (1000-200000, default 50000); content beyond is truncated"),
    }),
    execute: async (params) => unwrapResult(await readDocument(params)),
    domain: "rag",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "create_folder",
    description: "Create a new folder in OneDrive",
    parameters: z.object({
      parent_id: z.string().describe("Parent folder ID"),
      name: z.string().describe("Folder name"),
    }),
    execute: async (params) => unwrapResult(await createFolder(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "upload_file",
    description:
      "Upload TEXT content (plain text, markdown, CSV, JSON, HTML, XML) to OneDrive inline via this tool call. For binary files (docx, pdf, images, etc.), use get_upload_config (HTTP/SSE) or upload_file_from_path (stdio/local) — never base64-encode binary into this tool's content param. Max ~4 MB text.",
    parameters: z.object({
      path: z
        .string()
        .describe("Destination path in colon-path format (e.g., /me/drive/root:/Documents/file.txt:/content)"),
      content: z.string().describe("Text content to upload (UTF-8)"),
      content_type: z
        .string()
        .optional()
        .describe(
          "MIME type, default text/plain. Must be a text type (text/*, application/json, application/xml, application/javascript, *+json, *+xml). Binary types are rejected.",
        ),
      conflict_behavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns 409 on collision'),
    }),
    execute: async (params) => unwrapResult(await uploadFile(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "get_upload_config",
    description:
      "Get an authenticated upload URL + curl command for uploading files to OneDrive. Primary path for binary files or anything >1 MB. Pipe base64 file contents to the returned URL via POST; the server decodes and streams to Microsoft Graph (chunked session upload for >4 MB, up to 250 MB). Intermediate folders are auto-created. Note: response includes operational caveats in notes — read before executing the curl, especially on failure.",
    parameters: z.object({
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /me/drive/root:/Documents/file.docx:/content)",
        ),
      localFile: z
        .string()
        .optional()
        .describe("Local file path to include in the curl example. If omitted, a placeholder is used."),
      contentType: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflictBehavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns an error'),
    }),
    execute: async (params) => unwrapResult(await getUploadConfig(params)),
    domain: "files",
    readOnly: false,
  },
  {
    name: "upload_file_from_path",
    description:
      "Upload a local file to OneDrive by reading it from disk on the server. The file must exist on this machine's filesystem. If you generated the file in a cloud container (e.g., claude.ai), first use Desktop Commander's write_file to save it to the user's local filesystem, then call this tool with that local path. Supports files up to 250 MB (chunked session upload above 4 MB). Intermediate folders are auto-created.",
    parameters: z.object({
      local_path: z.string().describe("Absolute path to the local file to upload"),
      path: z
        .string()
        .describe(
          "Graph API destination path ending with :/content (e.g., /me/drive/root:/Documents/file.docx:/content)",
        ),
      content_type: z.string().optional().describe("MIME type override. Auto-detected from file extension if omitted."),
      conflict_behavior: z
        .enum(["rename", "replace", "fail"])
        .optional()
        .describe('Conflict behavior: "rename" (default), "replace" overwrites, "fail" returns an error'),
    }),
    execute: async (params) => unwrapResult(await uploadFileFromPath(params)),
    domain: "files",
    readOnly: false,
  },

  // === SharePoint Tools ===
  {
    name: "list_sites",
    description: "List SharePoint sites. Without a query, returns followed sites. With a query, searches all sites.",
    parameters: z.object({
      query: z.string().optional().describe("Search query to find sites (omit to list followed sites)"),
    }),
    execute: async (params) => unwrapResult(await listSites(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "get_site",
    description: "Get SharePoint site details",
    parameters: z.object({
      site_id: z.string().describe("Site ID (e.g., 'contoso.sharepoint.com,siteId,webId')"),
    }),
    execute: async (params) => unwrapResult(await getSite(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_site_drives",
    description: "List document libraries (drives) in a SharePoint site",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
    }),
    execute: async (params) => unwrapResult(await listSiteDrives(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "list_site_items",
    description: "List files and folders in a SharePoint site's document library",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
      drive_id: z.string().optional().describe("Drive ID (omit for default document library)"),
      folder_id: z.string().optional().describe("Folder ID (omit for root)"),
      folder_path: z.string().optional().describe("Folder path (e.g., 'General/Reports')"),
    }),
    execute: async (params) => unwrapResult(await listSiteItems(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
  {
    name: "search_site_files",
    description: "Search files within a SharePoint site",
    parameters: z.object({
      site_id: z.string().describe("Site ID"),
      query: z.string().describe("Search query"),
      drive_id: z.string().optional().describe("Drive ID (omit to search default document library)"),
    }),
    execute: async (params) => unwrapResult(await searchSiteFiles(params)),
    domain: "files",
    readOnly: true,
    annotations: { readOnlyHint: true },
  },
]
