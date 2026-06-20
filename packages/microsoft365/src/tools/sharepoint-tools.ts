import { UserError } from "fastmcp"
import type { Either } from "functype/either"
import { Left } from "functype/either"

import { getGraphClient } from "../client/graph-client"
import type { GraphDrive, GraphSite, ODataResponse } from "../types"
import { formatDriveItemList } from "../utils/formatters"

const requireClient = () => {
  const client = getGraphClient()
  if (client.isNone()) return null
  return client.orThrow()
}

const formatSiteSummary = (site: GraphSite): string =>
  `- **${site.displayName ?? site.name ?? "Untitled"}** (ID: ${site.id})${site.webUrl ? `\n  URL: ${site.webUrl}` : ""}${site.description ? `\n  ${site.description}` : ""}`

const formatSiteList = (sites: ReadonlyArray<GraphSite>): string =>
  sites.length === 0 ? "No sites found." : `# SharePoint Sites\n\n${sites.map(formatSiteSummary).join("\n")}`

const formatSiteDetail = (site: GraphSite): string =>
  `# ${site.displayName ?? site.name ?? "Untitled"}

## Details
- ID: ${site.id}
- Name: ${site.name ?? "N/A"}
- URL: ${site.webUrl ?? "N/A"}
- Description: ${site.description ?? "N/A"}
- Created: ${site.createdDateTime ?? "N/A"}
- Last Modified: ${site.lastModifiedDateTime ?? "N/A"}`

const formatDriveSummary = (drive: GraphDrive): string => {
  const quota = drive.quota
    ? ` (${formatSize(drive.quota.used ?? 0)} used of ${formatSize(drive.quota.total ?? 0)})`
    : ""
  return `- **${drive.name ?? "Untitled"}** (ID: ${drive.id}) - ${drive.driveType ?? "unknown"}${quota}`
}

const formatSize = (bytes: number): string => {
  if (bytes === 0) return "0 B"
  const units = ["B", "KB", "MB", "GB"]
  const i = Math.floor(Math.log(bytes) / Math.log(1024))
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${units[i]}`
}

const formatDriveList = (drives: ReadonlyArray<GraphDrive>): string =>
  drives.length === 0 ? "No drives found." : `# Document Libraries\n\n${drives.map(formatDriveSummary).join("\n")}`

export const listSites = async (params: { query?: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.query) {
    const result = await client.searchSites(params.query)
    return result
      .mapLeft((error) => new UserError(`Failed to search sites: ${error.message}`))
      .map((response) => formatSiteList((response as ODataResponse<never>).value))
  }

  const result = await client.listFollowedSites()
  return result
    .mapLeft((error) => new UserError(`Failed to list sites: ${error.message}`))
    .map((response) => formatSiteList((response as ODataResponse<never>).value))
}

export const getSite = async (params: { site_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.getSite(params.site_id)
  return result
    .mapLeft((error) => new UserError(`Failed to get site: ${error.message}`))
    .map((site) => formatSiteDetail(site as GraphSite))
}

export const listSiteDrives = async (params: { site_id: string }): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.listSiteDrives(params.site_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list drives: ${error.message}`))
    .map((response) => formatDriveList((response as ODataResponse<never>).value))
}

export const listSiteItems = async (params: {
  site_id: string
  drive_id?: string
  folder_id?: string
  folder_path?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  if (params.folder_path) {
    const result = await client.listSiteDriveItemsByPath(params.site_id, params.folder_path, params.drive_id)
    return result
      .mapLeft((error) => new UserError(`Failed to list site items: ${error.message}`))
      .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
  }

  const result = await client.listSiteDriveItems(params.site_id, params.drive_id, params.folder_id)
  return result
    .mapLeft((error) => new UserError(`Failed to list site items: ${error.message}`))
    .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
}

export const searchSiteFiles = async (params: {
  site_id: string
  query: string
  drive_id?: string
}): Promise<Either<UserError, string>> => {
  const client = requireClient()
  if (!client) return Left(new UserError("MS 365 client not initialized. Check authentication."))

  const result = await client.searchSiteFiles(params.site_id, params.query, params.drive_id)
  return result
    .mapLeft((error) => new UserError(`Failed to search site files: ${error.message}`))
    .map((response) => formatDriveItemList((response as ODataResponse<never>).value))
}
