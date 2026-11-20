import { ifMatchProblem } from "@sapientsai/ms-graph-core"
import type { Either } from "functype/either"
import { Left, Right } from "functype/either"
import { UserError } from "somamcp"

import { getGraphClient } from "../client/graph-client"
import type { GraphApiError } from "../types"
import { formatDriveItemDetail } from "../utils/formatters"

// Move, rename, copy and delete for OneDrive and SharePoint items. A drive_id selects a SharePoint
// document library (from list_site_drives); without one, the item is in the caller's own OneDrive.

const NOT_INITIALIZED = "MS 365 client not initialized. Check authentication."

export const ITEM_PRECONDITION_FAILED_MESSAGE =
  "precondition-failed: the item changed since it was read, so its eTag no longer matches. Nothing was changed. " +
  "Read it again with get_drive_item and retry with the new eTag."

// IDs go into the request path as they are. One holding a slash, "?", "#", "%" or a dot segment could
// point the PATCH or DELETE at a different Graph resource than the tool names. Real drive and item IDs
// use only letters, digits and ! _ . -, so anything else is refused before Graph is called.
const GRAPH_ID = /^[A-Za-z0-9!_.-]+$/

// JSON.stringify quotes the value and escapes any newline, so a refused ID can't add lines to the output.
const idProblem = (label: string, value: string | undefined): string | undefined => {
  if (value === undefined) return undefined
  if (!GRAPH_ID.test(value) || /^\.+$/.test(value)) {
    return `${label} ${JSON.stringify(value.slice(0, 80))} is not a valid ID. Use the ID from list_drive_items, get_drive_item or list_site_drives.`
  }
  return undefined
}

// "root" is a valid path alias, but never an item to delete, rename, move or copy.
const itemIdProblem = (value: string): string | undefined =>
  value.toLowerCase() === "root"
    ? "item_id root is the drive itself; pass the ID of a file or folder in it."
    : idProblem("item_id", value)

const firstProblem = (...problems: ReadonlyArray<string | undefined>): string | undefined =>
  problems.find((problem) => problem !== undefined)

// OneDrive and SharePoint refuse these characters in a name, and a slash would read as a path.
const INVALID_NAME_CHARS = /["*:<>?/\\|]/

const nameProblem = (name: string | undefined): string | undefined => {
  if (name === undefined) return undefined
  if (name.trim() === "") return "new_name is blank."
  // eslint-disable-next-line no-control-regex -- refusing control characters is the point
  if (/[\u0000-\u001f\u007f]/.test(name)) return "new_name can't contain control characters such as a line break."
  if (INVALID_NAME_CHARS.test(name)) return `new_name can't contain any of " * : < > ? / \\ |`
  return undefined
}

const failure = (action: string, error: GraphApiError): UserError =>
  error.status === 412
    ? new UserError(`Failed to ${action}: ${ITEM_PRECONDITION_FAILED_MESSAGE}`)
    : new UserError(`Failed to ${action}: ${error.message}`)

type Destination = { readonly folderId: string; readonly driveId?: string }

/**
 * Where an item goes. Graph refuses "root" in parentReference, so the root's real ID is looked up. For a
 * copy, parentReference should also name the drive; when the caller gave none (their own OneDrive),
 * the root lookup supplies it.
 */
const resolveDestination = async (
  folderId: string,
  driveId: string | undefined,
  needDriveId: boolean,
): Promise<Either<UserError, Destination>> => {
  if (folderId.toLowerCase() !== "root" && (driveId !== undefined || !needDriveId)) {
    return Right({ folderId, driveId })
  }
  const client = getGraphClient().orThrow()
  const root = await client.getDriveRoot(driveId)
  if (root.isLeft()) return Left(failure("find the drive's root folder", root.value as GraphApiError))
  const found = root.orThrow()
  return Right({
    folderId: folderId.toLowerCase() === "root" ? found.id : folderId,
    driveId: driveId ?? found.parentReference?.driveId,
  })
}

export const moveDriveItem = async (params: {
  item_id: string
  destination_folder_id: string
  drive_id?: string
  new_name?: string
  if_match?: string
}): Promise<Either<UserError, string>> => {
  const problem = firstProblem(
    itemIdProblem(params.item_id),
    idProblem("destination_folder_id", params.destination_folder_id),
    idProblem("drive_id", params.drive_id),
    nameProblem(params.new_name),
    ifMatchProblem(params.if_match),
  )
  if (problem) return Left(new UserError(problem))
  if (getGraphClient().isNone()) return Left(new UserError(NOT_INITIALIZED))

  const destination = await resolveDestination(params.destination_folder_id, params.drive_id, false)
  if (destination.isLeft()) return Left(destination.value as UserError)

  const changes = {
    parentReference: { id: destination.orThrow().folderId },
    ...(params.new_name !== undefined ? { name: params.new_name } : {}),
  }
  const result = await getGraphClient()
    .orThrow()
    .updateDriveItem(params.item_id, changes, params.drive_id, params.if_match)
  return result
    .mapLeft((error) => failure("move item", error))
    .map((item) => `Item moved.\n\n${formatDriveItemDetail(item)}`)
}

export const renameDriveItem = async (params: {
  item_id: string
  new_name: string
  drive_id?: string
  if_match?: string
}): Promise<Either<UserError, string>> => {
  const problem = firstProblem(
    itemIdProblem(params.item_id),
    idProblem("drive_id", params.drive_id),
    nameProblem(params.new_name),
    ifMatchProblem(params.if_match),
  )
  if (problem) return Left(new UserError(problem))
  const client = getGraphClient()
  if (client.isNone()) return Left(new UserError(NOT_INITIALIZED))

  const result = await client
    .orThrow()
    .updateDriveItem(params.item_id, { name: params.new_name }, params.drive_id, params.if_match)
  return result
    .mapLeft((error) => failure("rename item", error))
    .map((item) => `Item renamed.\n\n${formatDriveItemDetail(item)}`)
}

export const copyDriveItem = async (params: {
  item_id: string
  destination_folder_id: string
  drive_id?: string
  destination_drive_id?: string
  new_name?: string
  conflict_behavior?: "fail" | "rename"
}): Promise<Either<UserError, string>> => {
  const problem = firstProblem(
    itemIdProblem(params.item_id),
    idProblem("destination_folder_id", params.destination_folder_id),
    idProblem("drive_id", params.drive_id),
    idProblem("destination_drive_id", params.destination_drive_id),
    nameProblem(params.new_name),
  )
  if (problem) return Left(new UserError(problem))
  if (getGraphClient().isNone()) return Left(new UserError(NOT_INITIALIZED))

  // Without a destination drive the copy stays in the source item's drive.
  const resolved = await resolveDestination(
    params.destination_folder_id,
    params.destination_drive_id ?? params.drive_id,
    true,
  )
  if (resolved.isLeft()) return Left(resolved.value as UserError)
  const destination = resolved.orThrow()

  const body = {
    parentReference: {
      id: destination.folderId,
      ...(destination.driveId !== undefined ? { driveId: destination.driveId } : {}),
    },
    ...(params.new_name !== undefined ? { name: params.new_name } : {}),
  }
  const conflict = params.conflict_behavior === "rename" ? "rename" : undefined
  const result = await getGraphClient().orThrow().copyDriveItem(params.item_id, body, params.drive_id, conflict)
  return result
    .mapLeft((error) => failure("copy item", error))
    .map(
      () =>
        "Copy started. Graph copies in the background, so a large file or folder can take a while to appear. " +
        "List the destination folder with list_drive_items (or list_site_items) to check it has arrived. " +
        "If a same-named item was already there, the copy fails unless conflict_behavior is rename, and that " +
        "failure isn't reported here. Don't delete the original until the copy is confirmed in the destination.",
    )
}

export const deleteDriveItem = async (params: {
  item_id: string
  drive_id?: string
  if_match?: string
}): Promise<Either<UserError, string>> => {
  const problem = firstProblem(
    itemIdProblem(params.item_id),
    idProblem("drive_id", params.drive_id),
    ifMatchProblem(params.if_match),
  )
  if (problem) return Left(new UserError(problem))
  const client = getGraphClient()
  if (client.isNone()) return Left(new UserError(NOT_INITIALIZED))

  const result = await client.orThrow().deleteDriveItem(params.item_id, params.drive_id, params.if_match)
  return result
    .mapLeft((error) => failure("delete item", error))
    .map(
      () =>
        `Item ${params.item_id} moved to the recycle bin, with its contents if it is a folder. ` +
        "It can be restored from the OneDrive or SharePoint recycle bin.",
    )
}
