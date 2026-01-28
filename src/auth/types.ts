export type AuthMode = "interactive" | "clientCredentials"

export type ServerConfig = {
  clientId: string
  clientSecret: string
  tenantId: string
  baseUrl: string
  port: number
  scopes: string[]
  authMode: AuthMode
  appScopes: string[]
  apiKey?: string
}

export type AppOnlySession = {
  accessToken: string
  expiresAt: Date
  mode: "clientCredentials"
}

export type TokenResponse = {
  access_token: string
  expires_in: number
  token_type: string
}
