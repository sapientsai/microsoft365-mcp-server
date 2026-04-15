import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

import type { GraphApiError, GraphDriveItem } from "../types"

export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024 // 4 MB
export const MAX_UPLOAD_SIZE = 250 * 1024 * 1024 // 250 MB
const CHUNK_SIZE = 10 * 1024 * 1024 // 10 MB (must be multiple of 320 KiB)

const parseGraphError = async (response: Response): Promise<string> => {
  try {
    const body = (await response.json()) as { error?: { message?: string } }
    if (body?.error?.message) return body.error.message
  } catch {
    // fall through
  }
  return `${response.status} ${response.statusText}`
}

const toGraphApiError = async (response: Response, prefix: string): Promise<GraphApiError> => {
  const message = await parseGraphError(response)
  return {
    type: "api",
    message: `${prefix}: ${message}`,
    status: response.status,
  }
}

export const decodeBase64Upload = (rawBuffer: Buffer): Buffer =>
  Buffer.from(rawBuffer.toString("utf-8").replace(/\s/g, ""), "base64")

export const simpleUpload = async (
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  contentType: string,
  conflictBehavior: string,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const separator = path.includes("?") ? "&" : "?"
  const url = `${apiBase}${path}${separator}@microsoft.graph.conflictBehavior=${conflictBehavior}`

  try {
    const response = await fetch(url, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": contentType,
        "Content-Length": String(buffer.length),
      },
      body: new Uint8Array(buffer),
    })

    if (!response.ok) {
      return Left(await toGraphApiError(response, "Upload failed"))
    }

    const item = (await response.json()) as GraphDriveItem
    return Right(item)
  } catch (error) {
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error during upload: ${error instanceof Error ? error.message : String(error)}`,
    })
  }
}

export const sessionUpload = async (
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  conflictBehavior: string,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const sessionPath = path.replace(/:\/?content$/i, ":/createUploadSession")
  const sessionUrl = `${apiBase}${sessionPath}`

  try {
    const createResponse = await fetch(sessionUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": conflictBehavior },
      }),
    })

    if (!createResponse.ok) {
      return Left(await toGraphApiError(createResponse, "Failed to create upload session"))
    }

    const session = (await createResponse.json()) as { uploadUrl: string }
    return uploadChunks(session.uploadUrl, buffer, 0)
  } catch (error) {
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error creating upload session: ${error instanceof Error ? error.message : String(error)}`,
    })
  }
}

const uploadChunks = async (
  uploadUrl: string,
  buffer: Buffer,
  offset: number,
): Promise<Either<GraphApiError, GraphDriveItem>> => {
  const totalSize = buffer.length
  if (offset >= totalSize) {
    return Left<GraphApiError, GraphDriveItem>({
      type: "api",
      message: "Upload completed but no DriveItem response received",
    })
  }

  const end = Math.min(offset + CHUNK_SIZE, totalSize)
  const chunk = buffer.subarray(offset, end)

  try {
    const chunkResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: new Uint8Array(chunk),
    })

    if (!chunkResponse.ok) {
      await fetch(uploadUrl, { method: "DELETE" }).catch(() => {})
      return Left(await toGraphApiError(chunkResponse, `Upload chunk failed at byte ${offset}`))
    }

    if (chunkResponse.status === 200 || chunkResponse.status === 201) {
      const item = (await chunkResponse.json()) as GraphDriveItem
      return Right(item)
    }

    return uploadChunks(uploadUrl, buffer, offset + CHUNK_SIZE)
  } catch (error) {
    return Left<GraphApiError, GraphDriveItem>({
      type: "network",
      message: `Network error uploading chunk at byte ${offset}: ${error instanceof Error ? error.message : String(error)}`,
    })
  }
}
