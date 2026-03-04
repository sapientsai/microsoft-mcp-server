import { type Either, Left, Right, Try } from "functype"

import { type AuthError, authError } from "../errors.js"
import type { AppOnlySession, ServerConfig, TokenResponse } from "./types.js"

const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000 // 5 minutes before expiry

export type TokenManager = {
  getToken: () => Promise<Either<AuthError, string>>
  getSession: () => Promise<Either<AuthError, AppOnlySession>>
}

export function createTokenManager(config: Readonly<ServerConfig>): TokenManager {
  const state: { session: AppOnlySession | null } = { session: null }

  const fetchToken = async (): Promise<Either<AuthError, AppOnlySession>> => {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`

    const body = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      scope: config.appScopes.join(" "),
      grant_type: "client_credentials",
    })

    const response = await fetch(tokenUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    })

    if (!response.ok) {
      const errorText = await response.text()
      const errorMessage = Try(() => JSON.parse(errorText) as { error_description?: string; error?: string }).fold(
        () => `Token request failed: ${response.status} - ${errorText}`,
        (errorJson) => errorJson.error_description ?? errorJson.error ?? `Token request failed: ${response.status}`,
      )
      return Left(authError(errorMessage))
    }

    const data = (await response.json()) as TokenResponse

    return Right({
      accessToken: data.access_token,
      expiresAt: new Date(Date.now() + data.expires_in * 1000),
      mode: "clientCredentials" as const,
    })
  }

  const isTokenValid = (session: Readonly<AppOnlySession>): boolean => {
    const now = Date.now()
    const expiresAt = session.expiresAt.getTime()
    return now < expiresAt - TOKEN_REFRESH_BUFFER_MS
  }

  const getSession = async (): Promise<Either<AuthError, AppOnlySession>> => {
    if (state.session && isTokenValid(state.session)) {
      return Right(state.session)
    }
    const result = await fetchToken()
    return result.tap((session) => {
      state.session = session
    })
  }

  const getToken = async (): Promise<Either<AuthError, string>> => {
    const result = await getSession()
    return result.map((session) => session.accessToken)
  }

  return { getToken, getSession }
}
