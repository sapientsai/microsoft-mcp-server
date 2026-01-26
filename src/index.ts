import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"

import { AuthManager } from "./auth/index.js"
import { GraphClient } from "./client/index.js"
import { registerAuthTools, registerGraphTools } from "./tools/index.js"
import type { AuthMode, GraphApiVersion, ServerConfig } from "./types.js"
import { DEFAULT_SCOPES } from "./types.js"

export function createServerConfig(): ServerConfig {
  const clientId = process.env.AZURE_CLIENT_ID
  const tenantId = process.env.AZURE_TENANT_ID ?? "common"
  const clientSecret = process.env.AZURE_CLIENT_SECRET
  const accessToken = process.env.ACCESS_TOKEN

  if (!clientId && !accessToken) {
    throw new Error("AZURE_CLIENT_ID or ACCESS_TOKEN environment variable is required")
  }

  let authMode: AuthMode = "device_code"
  const envAuthMode = process.env.AUTH_MODE?.toLowerCase()
  if (envAuthMode === "client_credentials") {
    if (!clientSecret) {
      throw new Error("AZURE_CLIENT_SECRET is required for client_credentials mode")
    }
    authMode = "client_credentials"
  } else if (envAuthMode === "client_token" || accessToken) {
    authMode = "client_token"
  }

  const scopesEnv = process.env.GRAPH_SCOPES
  const scopes = scopesEnv ? scopesEnv.split(",").map((s) => s.trim()) : DEFAULT_SCOPES

  const graphApiVersion = (process.env.GRAPH_API_VERSION as GraphApiVersion) ?? "v1.0"

  return {
    clientId: clientId ?? "",
    tenantId,
    clientSecret,
    authMode,
    graphApiVersion,
    scopes,
  }
}

export function createServer(config: ServerConfig): McpServer {
  const server = new McpServer({
    name: "microsoft-graph-server",
    version: "1.0.0",
  })

  const authManager = new AuthManager(config)

  if (config.authMode === "client_token" && process.env.ACCESS_TOKEN) {
    authManager.setAccessToken(process.env.ACCESS_TOKEN)
  }

  const graphClient = new GraphClient(authManager, config.graphApiVersion)

  registerAuthTools(server, authManager)
  registerGraphTools(server, graphClient)

  return server
}

export async function runServer(): Promise<void> {
  const config = createServerConfig()
  const server = createServer(config)

  const transport = new StdioServerTransport()
  await server.connect(transport)

  console.error("Microsoft Graph MCP Server running on stdio")
}

export { AuthManager } from "./auth/index.js"
export { GraphClient } from "./client/index.js"
export * from "./types.js"
