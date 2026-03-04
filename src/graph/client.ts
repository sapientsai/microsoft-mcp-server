import { type Either, Left, Right } from "functype"

import { type GraphError, graphError } from "../errors.js"

export const GRAPH_BASE_URL = "https://graph.microsoft.com"
export const AZURE_BASE_URL = "https://management.azure.com"

export async function parseGraphError(response: Response): Promise<string> {
  const fallback = `HTTP ${response.status}: ${response.statusText}`
  try {
    const data = (await response.json()) as { error?: { message?: string; code?: string } }
    if (data.error?.message) {
      return data.error.code ? `${data.error.code}: ${data.error.message}` : data.error.message
    }
  } catch {
    // Not JSON — fall through
  }
  return fallback
}

export async function graphFetch(
  url: string,
  accessToken: string,
  init?: RequestInit,
): Promise<Either<GraphError, Response>> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...init?.headers,
    },
  })

  if (!response.ok) {
    const message = await parseGraphError(response)
    return Left(graphError(response.statusText, message, response.status))
  }

  return Right(response)
}
