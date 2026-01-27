# Microsoft Graph MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Graph API, enabling AI assistants to interact with Microsoft 365 services including users, mail, calendar, files, and more.

Built with [FastMCP](https://github.com/punkpeye/fastmcp) for seamless OAuth authentication.

## Features

- **Microsoft Graph API Access**: Execute any Graph API endpoint through a unified tool
- **Azure OAuth Authentication**: Automatic OAuth 2.0 flow handled by MCP client (Claude Desktop)
- **Full API Coverage**: Access Graph API v1.0 and beta endpoints
- **Azure Management API**: Optional support for Azure Resource Manager API
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
3. Add redirect URI: `http://localhost:8080/oauth/callback`
4. Create a client secret
5. Grant API permissions for Microsoft Graph (e.g., User.Read, Mail.Read)

### 2. Configure Environment

Create a `.env` file:

```bash
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=common  # or specific tenant ID

# Server Configuration
BASE_URL=http://localhost:8080
PORT=8080

# Transport: httpStream (default) or stdio
TRANSPORT_TYPE=httpStream

# Optional: Custom scopes
# GRAPH_SCOPES=openid,profile,email,User.Read,Mail.Read
```

### 3. Run the Server

```bash
npx microsoft-mcp-server
```

The server starts on `http://localhost:8080` with OAuth endpoint at `/oauth/callback`.

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

Check current authentication status. Returns authentication status, scopes, and user principal name.

## Authentication

Authentication is handled automatically by the MCP client (Claude Desktop) via OAuth 2.0:

1. When you first use the `microsoft_graph` tool, Claude Desktop will prompt for authentication
2. You'll be redirected to Microsoft's login page
3. After successful login, the token is managed by FastMCP
4. Subsequent requests use the cached token automatically

### Required Azure App Permissions

For delegated (user) access, add these Microsoft Graph API permissions:

- `User.Read` - Read user profile
- `Mail.Read` - Read user mail (optional)
- `Calendars.Read` - Read user calendars (optional)
- `Files.Read` - Read user files (optional)
- `Sites.Read.All` - Read SharePoint sites (optional)

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
