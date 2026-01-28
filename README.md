# Microsoft Graph MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Graph API, enabling AI assistants to interact with Microsoft 365 services including users, mail, calendar, files, and more.

Built with [FastMCP](https://github.com/punkpeye/fastmcp) for seamless OAuth authentication.

## Features

- **Microsoft Graph API Access**: Execute any Graph API endpoint through a unified tool
- **Dual Authentication Modes**:
  - **Interactive** (default): OAuth 2.0 authorization code flow with user login
  - **Client Credentials**: App-only authentication for headless/server deployments
- **Full API Coverage**: Access Graph API v1.0 and beta endpoints
- **Azure Management API**: Optional support for Azure Resource Manager API
- **API Key Protection**: Optional endpoint security for production deployments
- **HTTP & stdio transports**: Run as HTTP server or stdio-based MCP

## Installation

```bash
npm install microsoft-mcp-server
# or
pnpm add microsoft-mcp-server
```

## Quick Start

### 1. Create Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Create a new registration
3. Add redirect URI: `http://localhost:8080/oauth/callback` (for interactive mode)
4. Create a client secret
5. Grant API permissions for Microsoft Graph (see [Permissions](#azure-app-permissions) below)

### 2. Configure Environment

Create a `.env` file:

```bash
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=common  # or specific tenant ID

# Auth mode: 'interactive' (default) or 'clientCredentials'
AZURE_AUTH_MODE=interactive

# Server Configuration
BASE_URL=http://localhost:8080
PORT=8080

# Transport: httpStream (default) or stdio
TRANSPORT_TYPE=httpStream

# Optional: Custom scopes for interactive mode
# GRAPH_SCOPES=openid,profile,email,User.Read,Mail.Read

# Optional: API key protection
# MCP_API_KEY=your-secret-key
```

### 3. Run the Server

```bash
npx microsoft-mcp-server
```

The server starts on `http://localhost:8080` with OAuth endpoint at `/oauth/callback`.

## Authentication Modes

### Interactive Mode (Default)

User-based authentication via OAuth 2.0 authorization code flow. Best for:

- Desktop applications
- Development/testing
- Scenarios requiring user-specific permissions

```bash
AZURE_AUTH_MODE=interactive
AZURE_TENANT_ID=common  # or specific tenant
```

When you first use the `microsoft_graph` tool, the MCP client (Claude Desktop) prompts for login. After successful authentication, the token is cached automatically.

### Client Credentials Mode

App-only authentication for headless/server deployments. Best for:

- Background services
- Automated workflows
- Server-to-server communication
- CI/CD pipelines

```bash
AZURE_AUTH_MODE=clientCredentials
AZURE_TENANT_ID=your-specific-tenant-id  # Required: cannot use "common"
AZURE_CLIENT_SECRET=your-client-secret   # Required
GRAPH_APP_SCOPES=https://graph.microsoft.com/.default
```

**Important**: Client credentials mode requires:

- A specific tenant ID (not "common")
- A client secret
- **Application permissions** (not Delegated) configured in Azure
- **Admin consent** granted by a tenant administrator

## Usage

### With Claude Desktop (HTTP Mode)

Add to your Claude Desktop config:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "url": "http://localhost:8080/mcp"
    }
  }
}
```

### With Claude Code CLI (stdio Mode)

Add to your project's `.mcp.json`:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["microsoft-mcp-server"],
      "env": {
        "TRANSPORT_TYPE": "stdio",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Client Credentials Example

For headless server deployments:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["microsoft-mcp-server"],
      "env": {
        "TRANSPORT_TYPE": "stdio",
        "AZURE_AUTH_MODE": "clientCredentials",
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## Available Tools

### `microsoft_graph`

Execute Microsoft Graph API requests.

**Parameters:**

| Parameter     | Required | Description                                               |
| ------------- | -------- | --------------------------------------------------------- |
| `path`        | Yes      | API endpoint path (e.g., `/me`, `/users`, `/me/messages`) |
| `method`      | No       | HTTP method: GET, POST, PUT, PATCH, DELETE (default: GET) |
| `apiVersion`  | No       | Graph API version: v1.0, beta (default: v1.0)             |
| `apiType`     | No       | API type: graph, azure (default: graph)                   |
| `queryParams` | No       | OData query parameters ($select, $filter, $top, etc.)     |
| `body`        | No       | Request body for POST/PUT/PATCH operations                |

**Example prompts to Claude:**

- "Get my profile information from Microsoft Graph"
- "Show me my last 10 emails"
- "List all users in my organization"
- "Create a calendar event for tomorrow at 2pm titled 'Team Sync'"
- "Search for files containing 'budget' in my OneDrive"

### `get_auth_status`

Check current authentication status. Returns:

- Authentication status
- Auth mode (interactive or clientCredentials)
- Scopes and user principal name (interactive mode)
- Token expiry time (client credentials mode)

## Azure App Permissions

### For Interactive Mode (Delegated Permissions)

Add these Microsoft Graph API **Delegated** permissions:

- `User.Read` - Read user profile
- `Mail.Read` - Read user mail (optional)
- `Calendars.Read` - Read user calendars (optional)
- `Files.Read` - Read user files (optional)
- `Sites.Read.All` - Read SharePoint sites (optional)

### For Client Credentials Mode (Application Permissions)

Add these Microsoft Graph API **Application** permissions:

- `User.Read.All` - Read all users' profiles
- `Mail.Read` - Read mail in all mailboxes (optional)
- `Calendars.Read` - Read calendars in all mailboxes (optional)
- `Files.Read.All` - Read all files (optional)
- `Sites.Read.All` - Read all SharePoint sites (optional)

**Important**: Application permissions require **admin consent**. A tenant administrator must grant consent in the Azure portal.

## API Key Protection

For production deployments, you can protect the MCP endpoint with an API key:

```bash
MCP_API_KEY=your-secret-api-key
```

When set, all requests must include the `Authorization: Bearer <key>` header.

## Environment Variables

| Variable              | Required    | Default                                | Description                                                 |
| --------------------- | ----------- | -------------------------------------- | ----------------------------------------------------------- |
| `AZURE_CLIENT_ID`     | Yes         | -                                      | Azure app registration client ID                            |
| `AZURE_CLIENT_SECRET` | Conditional | -                                      | Required for client credentials mode                        |
| `AZURE_TENANT_ID`     | No          | `common`                               | Tenant ID (specific tenant required for client credentials) |
| `AZURE_AUTH_MODE`     | No          | `interactive`                          | Auth mode: `interactive` or `clientCredentials`             |
| `BASE_URL`            | No          | `http://localhost:8080`                | Server URL for OAuth callback                               |
| `PORT`                | No          | `8080`                                 | Server port                                                 |
| `TRANSPORT_TYPE`      | No          | `httpStream`                           | Transport: `httpStream` or `stdio`                          |
| `GRAPH_SCOPES`        | No          | See below                              | Delegated scopes for interactive mode                       |
| `GRAPH_APP_SCOPES`    | No          | `https://graph.microsoft.com/.default` | App scopes for client credentials                           |
| `MCP_API_KEY`         | No          | -                                      | API key for endpoint protection                             |

**Default GRAPH_SCOPES**: `openid,profile,email,User.Read,Mail.Read,Calendars.Read,Files.Read,Sites.Read.All`

## Development

```bash
pnpm install          # Install dependencies
pnpm dev              # Development with watch
pnpm test             # Run tests
pnpm build            # Build for production
pnpm validate         # Format + lint + test + build
```

## Architecture

This server is built with [FastMCP](https://github.com/punkpeye/fastmcp), which provides:

- Automatic OAuth 2.0 flow with Azure AD
- HTTP streaming and SSE transport support
- Session management
- Health check endpoints

## License

MIT
