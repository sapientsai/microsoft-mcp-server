# Microsoft MCP Server: Client Credentials Flow Support

> **Purpose**: Add app-only (client credentials) authentication to microsoft-mcp-server for headless/server deployments.

| Field            | Value      |
| ---------------- | ---------- |
| **Version**      | 1.0-Draft  |
| **Status**       | Draft      |
| **Last Updated** | 2026-01-28 |
| **Author**       | SapientsAI |

---

## Problem Statement

The current microsoft-mcp-server only supports interactive OAuth 2.0 (authorization code flow), which requires a user to authenticate via browser. This prevents deployment as a headless service for:

- Server-to-server integrations
- Automated agent workflows
- Multi-tenant SaaS deployments

## Solution Overview

Add a new `authMode` configuration option that supports:

1. **`interactive`** (current behavior) - User authenticates via browser OAuth flow
2. **`clientCredentials`** (new) - App-only auth using Azure AD client credentials grant

Also add optional API key authentication to protect the MCP endpoint itself.

## Configuration Changes

### New Environment Variables

```bash
# MCP Endpoint Security (NEW)
MCP_API_KEY=your-secret-key  # Optional - if set, requires ?api_key= or Authorization header

# Auth mode selector (NEW)
AZURE_AUTH_MODE=clientCredentials  # or 'interactive' (default)

# Existing (used by both modes)
AZURE_CLIENT_ID=your-app-id
AZURE_CLIENT_SECRET=your-secret
AZURE_TENANT_ID=your-tenant-id  # Required for clientCredentials, 'common' not allowed
PORT=8080
TRANSPORT_TYPE=httpStream

# Interactive mode only
BASE_URL=https://your-server.com  # For OAuth callback
GRAPH_SCOPES=User.Read,Mail.Read  # Delegated permissions

# Client credentials mode only (NEW)
GRAPH_APP_SCOPES=https://graph.microsoft.com/.default  # App permissions
```

### ServerConfig Type Extension

```typescript
type AuthMode = "interactive" | "clientCredentials"

type ServerConfig = {
  // Existing fields
  clientId: string
  clientSecret: string
  tenantId: string
  baseUrl: string
  port: number
  scopes: string[]

  // New fields
  authMode: AuthMode
  appScopes?: string[] // For client credentials mode
  apiKey?: string // For MCP endpoint security
}
```

## Implementation Design

### 0. MCP Endpoint Security (New)

Add API key validation middleware to protect the MCP endpoint from unauthorized access.

Create `src/middleware/api-key.ts`:

```typescript
export function createApiKeyMiddleware(apiKey: string | undefined) {
  // If no API key configured, allow all requests
  if (!apiKey) {
    return (req: Request, next: () => Promise<Response>) => next()
  }

  return (req: Request, next: () => Promise<Response>) => {
    const url = new URL(req.url)

    // Check query param: ?api_key=xxx
    const queryKey = url.searchParams.get("api_key")

    // Check Authorization header: Bearer xxx or ApiKey xxx
    const authHeader = req.headers.get("Authorization")
    const headerKey = authHeader?.replace(/^(Bearer|ApiKey)\s+/i, "")

    const providedKey = queryKey || headerKey

    if (providedKey !== apiKey) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401,
        headers: { "Content-Type": "application/json" },
      })
    }

    return next()
  }
}
```

**Usage in server:**

```typescript
// In createServer()
const apiKeyMiddleware = createApiKeyMiddleware(process.env.MCP_API_KEY)

// Apply to HTTP transport
server.use(apiKeyMiddleware)
```

**Client configuration:**

```json
{
  "mcpServers": {
    "microsoft": {
      "type": "http",
      "url": "https://your-server.com/mcp?api_key=your-secret-key"
    }
  }
}
```

Or with header:

```json
{
  "mcpServers": {
    "microsoft": {
      "type": "http",
      "url": "https://your-server.com/mcp",
      "headers": {
        "Authorization": "ApiKey your-secret-key"
      }
    }
  }
}
```

### 1. Token Manager (New Module)

Create `src/auth/token-manager.ts`:

```typescript
interface TokenCache {
  accessToken: string
  expiresAt: Date
  scopes: string[]
}

class ClientCredentialsTokenManager {
  private cache: TokenCache | null = null
  private refreshBuffer = 5 * 60 * 1000 // 5 minutes before expiry

  constructor(
    private config: {
      clientId: string
      clientSecret: string
      tenantId: string
      scopes: string[]
    },
  ) {}

  async getToken(): Promise<string> {
    if (this.isTokenValid()) {
      return this.cache!.accessToken
    }
    return this.refreshToken()
  }

  private isTokenValid(): boolean {
    if (!this.cache) return false
    return Date.now() < this.cache.expiresAt.getTime() - this.refreshBuffer
  }

  private async refreshToken(): Promise<string> {
    const tokenEndpoint = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`

    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: this.config.clientId,
        client_secret: this.config.clientSecret,
        scope: this.config.scopes.join(" "),
      }),
    })

    if (!response.ok) {
      const error = await response.text()
      throw new Error(`Token request failed: ${error}`)
    }

    const data = await response.json()

    this.cache = {
      accessToken: data.access_token,
      expiresAt: new Date(Date.now() + data.expires_in * 1000),
      scopes: this.config.scopes,
    }

    return this.cache.accessToken
  }
}
```

### 2. Auth Provider Abstraction

Create `src/auth/providers.ts`:

```typescript
interface AuthProvider {
  getToken(): Promise<string>
  getAuthStatus(): { authenticated: boolean; mode: AuthMode; scopes?: string[] }
}

// For client credentials mode
class AppOnlyAuthProvider implements AuthProvider {
  private tokenManager: ClientCredentialsTokenManager

  constructor(config: ServerConfig) {
    this.tokenManager = new ClientCredentialsTokenManager({
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      tenantId: config.tenantId,
      scopes: config.appScopes || ["https://graph.microsoft.com/.default"],
    })
  }

  async getToken(): Promise<string> {
    return this.tokenManager.getToken()
  }

  getAuthStatus() {
    return { authenticated: true, mode: "clientCredentials" as AuthMode }
  }
}

// For interactive mode (wraps FastMCP session)
class InteractiveAuthProvider implements AuthProvider {
  constructor(private session: OAuthSession | undefined) {}

  async getToken(): Promise<string> {
    if (!this.session?.accessToken) {
      throw new Error("Not authenticated. Please sign in first.")
    }
    return this.session.accessToken
  }

  getAuthStatus() {
    return {
      authenticated: !!this.session?.accessToken,
      mode: "interactive" as AuthMode,
      scopes: this.session?.scopes,
    }
  }
}
```

### 3. Server Factory Changes

Modify `src/index.ts`:

```typescript
export function createServer(config: ServerConfig) {
  // Shared token manager for client credentials mode
  let appAuthProvider: AppOnlyAuthProvider | null = null

  if (config.authMode === "clientCredentials") {
    appAuthProvider = new AppOnlyAuthProvider(config)
  }

  const server = new FastMCP({
    name: "Microsoft Graph MCP Server",
    version: "1.0.0",
    // Only use AzureProvider for interactive mode
    ...(config.authMode === "interactive" && {
      auth: new AzureProvider({
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        baseUrl: config.baseUrl,
        tenantId: config.tenantId,
        scopes: config.scopes,
      }),
    }),
  })

  // Tool registration with auth mode awareness
  server.addTool({
    name: "microsoft_graph",
    // ... schema ...
    execute: async (args, { session }) => {
      // Get token based on auth mode
      const token =
        config.authMode === "clientCredentials"
          ? await appAuthProvider!.getToken()
          : await new InteractiveAuthProvider(session as OAuthSession).getToken()

      // Make Graph API call with token
      const response = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${token}` },
      })
      // ...
    },
  })

  return server
}
```

### 4. Configuration Validation

Add to `createConfig()`:

```typescript
export function createConfig(): ServerConfig {
  const authMode = (process.env.AZURE_AUTH_MODE || "interactive") as AuthMode

  // Validate auth mode
  if (!["interactive", "clientCredentials"].includes(authMode)) {
    throw new Error(`Invalid AZURE_AUTH_MODE: ${authMode}`)
  }

  // Client credentials requires specific tenant (not 'common')
  if (authMode === "clientCredentials" && (!process.env.AZURE_TENANT_ID || process.env.AZURE_TENANT_ID === "common")) {
    throw new Error("Client credentials mode requires a specific AZURE_TENANT_ID")
  }

  // Client credentials doesn't need BASE_URL
  const baseUrl = authMode === "interactive" ? process.env.BASE_URL || "http://localhost:8080" : ""

  return {
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    tenantId: process.env.AZURE_TENANT_ID || "common",
    baseUrl,
    port: parseInt(process.env.PORT || "8080"),
    scopes: parseScopes(process.env.GRAPH_SCOPES),
    authMode,
    apiKey: process.env.MCP_API_KEY,
    appScopes:
      authMode === "clientCredentials"
        ? parseScopes(process.env.GRAPH_APP_SCOPES || "https://graph.microsoft.com/.default")
        : undefined,
  }
}
```

## Azure AD Setup Differences

### Interactive Mode (Current)

- **Permission Type**: Delegated permissions
- **Admin Consent**: Per-user or admin consent
- **Scopes**: `User.Read`, `Mail.Read`, `Files.Read`, etc.

### Client Credentials Mode (New)

- **Permission Type**: Application permissions
- **Admin Consent**: Required (tenant admin must grant)
- **Scopes**: `https://graph.microsoft.com/.default` (gets all granted app permissions)
- **Sites.Selected**: Requires SharePoint admin to grant site-specific access

## File Changes Summary

| File                        | Change Type | Description                          |
| --------------------------- | ----------- | ------------------------------------ |
| `src/middleware/api-key.ts` | New         | MCP endpoint API key validation      |
| `src/auth/token-manager.ts` | New         | Client credentials token caching     |
| `src/auth/providers.ts`     | New         | Auth provider abstraction            |
| `src/index.ts`              | Modify      | Add auth mode branching + middleware |
| `.env.example`              | Modify      | Add new env vars                     |
| `README.md`                 | Modify      | Document both modes + API key        |
| `test/auth.spec.ts`         | New         | Auth provider tests                  |
| `test/api-key.spec.ts`      | New         | API key middleware tests             |

## Testing Plan

### Unit Tests

- Token manager caching behavior
- Token refresh before expiry
- Error handling for failed token requests
- Config validation for each mode
- API key middleware validation

### Integration Tests

- Client credentials flow against Azure AD
- Graph API call with app-only token
- Mode switching based on env vars
- API key rejection/acceptance

### Manual Tests

- Deploy with client credentials
- Verify SharePoint access with Sites.Selected
- Test API key protection

## Migration Path

1. **No breaking changes** - Default `authMode` is `interactive` (current behavior)
2. **Opt-in** - Set `AZURE_AUTH_MODE=clientCredentials` to use new flow
3. **Existing deployments** - No changes required unless switching modes

## Security Considerations

- Client secrets already stored in environment (no change)
- App-only tokens have no user context (expected for service accounts)
- Sites.Selected permission limits SharePoint access to explicitly granted sites
- Token caching is in-memory only (no persistence across restarts)
- API key adds endpoint-level protection for public deployments

## Timeline Estimate

| Phase     | Tasks                        | Estimate       |
| --------- | ---------------------------- | -------------- |
| 1         | API key middleware + tests   | 1-2 hours      |
| 2         | Token manager + tests        | 2-3 hours      |
| 3         | Auth provider abstraction    | 1-2 hours      |
| 4         | Server factory changes       | 1-2 hours      |
| 5         | Documentation + .env.example | 1 hour         |
| 6         | Integration testing          | 2-3 hours      |
| **Total** |                              | **8-13 hours** |

## Open Questions

1. Should we support certificate-based authentication in addition to client secrets?
2. Should token cache support Redis/external storage for multi-instance deployments?
3. Should we add a health check endpoint that validates token acquisition?
