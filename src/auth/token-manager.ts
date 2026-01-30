import type { AppOnlySession, ServerConfig, TokenResponse } from "./types.js"

const TOKEN_REFRESH_BUFFER_MS = 5 * 60 * 1000 // 5 minutes before expiry

export type TokenManager = {
  getToken: () => Promise<string>
  getSession: () => Promise<AppOnlySession>
}

export function createTokenManager(config: Readonly<ServerConfig>): TokenManager {
  // Use closure-based caching with explicit state type
  const state: { session: AppOnlySession | null } = { session: null }

  const fetchToken = async (): Promise<AppOnlySession> => {
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
      const errorMessage = (() => {
        try {
          const errorJson = JSON.parse(errorText) as { error_description?: string; error?: string }
          return errorJson.error_description ?? errorJson.error ?? `Token request failed: ${response.status}`
        } catch {
          return `Token request failed: ${response.status} - ${errorText}`
        }
      })()
      throw new Error(errorMessage)
    }

    const data = (await response.json()) as TokenResponse

    return {
      accessToken: data.access_token,
      expiresAt: new Date(Date.now() + data.expires_in * 1000),
      mode: "clientCredentials",
    }
  }

  const isTokenValid = (session: Readonly<AppOnlySession>): boolean => {
    const now = Date.now()
    const expiresAt = session.expiresAt.getTime()
    return now < expiresAt - TOKEN_REFRESH_BUFFER_MS
  }

  const getSession = async (): Promise<AppOnlySession> => {
    if (state.session && isTokenValid(state.session)) {
      return state.session
    }
    state.session = await fetchToken()
    return state.session
  }

  const getToken = async (): Promise<string> => {
    const session = await getSession()
    return session.accessToken
  }

  return { getToken, getSession }
}
