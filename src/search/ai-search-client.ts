import { type Either, Left, Right } from "functype"

import { type GraphError, graphError } from "../errors.js"

export const AI_SEARCH_API_VERSION = "2025-09-01"

export type AiSearchConfig = {
  readonly endpoint: string
  readonly apiKey: string
  readonly indexName: string
  readonly semanticConfiguration?: string
  readonly vectorFields?: string
  readonly selectFields?: string
}

export async function parseAiSearchError(response: Response): Promise<string> {
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

export async function aiSearchFetch(
  url: string,
  apiKey: string,
  init?: RequestInit,
): Promise<Either<GraphError, Response>> {
  const response = await fetch(url, {
    ...init,
    headers: {
      "api-key": apiKey,
      ...init?.headers,
    },
  })

  if (!response.ok) {
    const message = await parseAiSearchError(response)
    return Left(graphError(response.statusText, message, response.status))
  }

  return Right(response)
}
