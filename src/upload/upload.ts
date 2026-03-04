import { type Either, Left, Right } from "functype"

import { type GraphError, graphError } from "../errors.js"
import { parseGraphError } from "../graph/client.js"

export const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024 // 4 MB
export const MAX_UPLOAD_SIZE = 250 * 1024 * 1024 // 250 MB
const CHUNK_SIZE = 10 * 1024 * 1024 // 10 MB (must be multiple of 320 KiB)

export type DriveItemResponse = {
  id: string
  name: string
  size: number
  webUrl: string
  createdDateTime?: string
  lastModifiedDateTime?: string
  file?: { mimeType: string }
  parentReference?: { driveId: string; path: string }
}

export function decodeBase64Upload(rawBuffer: Buffer): Buffer {
  return Buffer.from(rawBuffer.toString("utf-8").replace(/\s/g, ""), "base64")
}

export async function simpleUpload(
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  contentType: string,
  conflictBehavior: string,
): Promise<Either<GraphError, DriveItemResponse>> {
  const separator = path.includes("?") ? "&" : "?"
  const url = `${apiBase}${path}${separator}@microsoft.graph.conflictBehavior=${conflictBehavior}`

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
    const message = await parseGraphError(response)
    return Left(graphError("upload_failed", message, response.status))
  }

  return Right((await response.json()) as DriveItemResponse)
}

export async function sessionUpload(
  apiBase: string,
  path: string,
  accessToken: string,
  buffer: Buffer,
  conflictBehavior: string,
): Promise<Either<GraphError, DriveItemResponse>> {
  // The path ends with :/content — replace :/content with :/createUploadSession
  const sessionPath = path.replace(/:\/?content$/i, ":/createUploadSession")
  const sessionUrl = `${apiBase}${sessionPath}`

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
    const message = await parseGraphError(createResponse)
    return Left(
      graphError("session_create_failed", `Failed to create upload session: ${message}`, createResponse.status),
    )
  }

  const session = (await createResponse.json()) as { uploadUrl: string }

  return uploadChunks(session.uploadUrl, buffer, 0)
}

async function uploadChunks(
  uploadUrl: string,
  buffer: Buffer,
  offset: number,
): Promise<Either<GraphError, DriveItemResponse>> {
  const totalSize = buffer.length
  if (offset >= totalSize) {
    return Left(graphError("upload_incomplete", "Upload completed but no DriveItem response received", 0))
  }

  const end = Math.min(offset + CHUNK_SIZE, totalSize)
  const chunk = buffer.subarray(offset, end)

  const chunkResponse = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": String(chunk.length),
      "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
    },
    body: new Uint8Array(chunk),
  })

  if (!chunkResponse.ok) {
    // Cancel the upload session on failure
    await fetch(uploadUrl, { method: "DELETE" }).catch(() => {})
    const message = await parseGraphError(chunkResponse)
    return Left(
      graphError("chunk_upload_failed", `Upload chunk failed at byte ${offset}: ${message}`, chunkResponse.status),
    )
  }

  // The final chunk returns the DriveItem; intermediate chunks return 202
  if (chunkResponse.status === 200 || chunkResponse.status === 201) {
    return Right((await chunkResponse.json()) as DriveItemResponse)
  }

  return uploadChunks(uploadUrl, buffer, offset + CHUNK_SIZE)
}
