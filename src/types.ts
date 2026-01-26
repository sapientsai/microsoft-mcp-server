export type AuthMode = "device_code" | "client_credentials" | "client_token"

export type HttpMethod = "GET" | "POST" | "PUT" | "PATCH" | "DELETE"

export type ApiType = "graph" | "azure"

export type GraphApiVersion = "v1.0" | "beta"

export type AuthStatus = {
  authenticated: boolean
  mode: AuthMode
  expiresOn?: string
  scopes?: string[]
  account?: string
}

export type TokenInfo = {
  accessToken: string
  expiresOn: Date
  scopes?: string[]
  account?: string
}

export type GraphApiRequest = {
  apiType?: ApiType
  path: string
  method?: HttpMethod
  apiVersion?: GraphApiVersion
  queryParams?: Record<string, string>
  body?: unknown
}

export type GraphApiResponse = {
  status: number
  data: unknown
  headers?: Record<string, string>
}

export type DeviceCodeInfo = {
  userCode: string
  verificationUri: string
  message: string
  expiresOn: Date
}

export type ServerConfig = {
  clientId: string
  tenantId: string
  clientSecret?: string
  authMode: AuthMode
  graphApiVersion: GraphApiVersion
  scopes: string[]
}

export const DEFAULT_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Mail.Read",
  "Calendars.Read",
  "Files.Read",
  "Sites.Read.All",
]

export const GRAPH_BASE_URL = "https://graph.microsoft.com"
export const AZURE_BASE_URL = "https://management.azure.com"
