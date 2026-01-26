# Microsoft Graph MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Graph API, enabling AI assistants to interact with Microsoft 365 services including users, mail, calendar, files, and more.

## Features

- **Microsoft Graph API Access**: Execute any Graph API endpoint through a unified tool
- **Multiple Authentication Modes**:
  - **Device Code Flow**: Interactive browser-based authentication
  - **Client Credentials**: App-only authentication with client secret
  - **Client Token**: Manual token injection for flexibility
- **Full API Coverage**: Access Graph API v1.0 and beta endpoints
- **Azure Management API**: Optional support for Azure Resource Manager API

## Installation

```bash
npm install microsoft-mcp-server
# or
pnpm add microsoft-mcp-server
```

## Prerequisites

### Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Click "New registration"
3. Name your app (e.g., "MCP Graph Server")
4. Select "Accounts in any organizational directory and personal Microsoft accounts"
5. Set Redirect URI to `https://login.microsoftonline.com/common/oauth2/nativeclient` (for device code flow)
6. Click "Register"

### Configure API Permissions

1. In your app registration, go to "API permissions"
2. Click "Add a permission" > "Microsoft Graph"
3. Add the permissions you need:
   - **Delegated** (for device_code): User.Read, Mail.Read, Calendars.Read, etc.
   - **Application** (for client_credentials): User.Read.All, Mail.Read, etc.
4. Grant admin consent if required

### Enable Public Client (for Device Code Flow)

1. Go to "Authentication"
2. Under "Advanced settings", set "Allow public client flows" to "Yes"
3. Save

## Configuration

Create a `.env` file or set environment variables:

```bash
# Required: Azure App Registration
AZURE_CLIENT_ID=your-client-id
AZURE_TENANT_ID=common  # or specific tenant ID

# Authentication Mode (default: device_code)
AUTH_MODE=device_code  # device_code | client_credentials | client_token

# For client_credentials mode
AZURE_CLIENT_SECRET=your-client-secret

# For client_token mode
ACCESS_TOKEN=your-access-token

# Optional: Graph API Configuration
GRAPH_API_VERSION=v1.0  # v1.0 | beta
GRAPH_SCOPES=User.Read,Mail.Read,Calendars.Read
```

## Usage

### With Claude Desktop

Add to your Claude Desktop config (`~/.config/claude/claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["microsoft-mcp-server"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "common",
        "AUTH_MODE": "device_code"
      }
    }
  }
}
```

### Standalone

```bash
# Run the server
npx microsoft-mcp-server

# Or if installed globally
microsoft-mcp-server
```

## Available Tools

### `microsoft_graph`

Execute Microsoft Graph API requests.

**Parameters:**

- `path` (required): API endpoint path (e.g., `/me`, `/users`, `/me/messages`)
- `method`: HTTP method (GET, POST, PUT, PATCH, DELETE) - default: GET
- `apiVersion`: Graph API version (v1.0, beta) - default: v1.0
- `apiType`: API type (graph, azure) - default: graph
- `queryParams`: OData query parameters ($select, $filter, $top, etc.)
- `body`: Request body for POST/PUT/PATCH operations

**Examples:**

```
# Get current user
microsoft_graph(path="/me")

# Get user's emails with filtering
microsoft_graph(
  path="/me/messages",
  queryParams={"$select": "subject,from,receivedDateTime", "$top": "10"}
)

# Create a calendar event
microsoft_graph(
  path="/me/events",
  method="POST",
  body={
    "subject": "Team Meeting",
    "start": {"dateTime": "2024-01-15T10:00:00", "timeZone": "UTC"},
    "end": {"dateTime": "2024-01-15T11:00:00", "timeZone": "UTC"}
  }
)
```

### `get_auth_status`

Check current authentication status.

**Returns:** Authentication mode, token validity, scopes, and account info.

### `set_access_token`

Manually set an access token for authentication.

**Parameters:**

- `accessToken` (required): Bearer access token
- `expiresOn`: ISO datetime when token expires (default: 1 hour from now)

### `sign_in`

Initiate device code authentication flow. Displays a code and URL for browser sign-in.

### `sign_out`

Clear all cached authentication tokens.

## Authentication Modes

### Device Code Flow (Recommended for Interactive Use)

1. Set `AUTH_MODE=device_code`
2. Call the `sign_in` tool
3. Visit the displayed URL and enter the code
4. Complete authentication in browser
5. Token is cached for subsequent requests

### Client Credentials (App-Only)

1. Set `AUTH_MODE=client_credentials`
2. Set `AZURE_CLIENT_SECRET`
3. Requires application permissions (not delegated)
4. No user interaction required

### Client Token (Manual)

1. Set `AUTH_MODE=client_token`
2. Either set `ACCESS_TOKEN` env var or use `set_access_token` tool
3. Useful for testing or when tokens are obtained externally

## Development

```bash
# Install dependencies
pnpm install

# Development with watch
pnpm dev

# Run tests
pnpm test

# Build for production
pnpm build

# Validate (format + lint + test + build)
pnpm validate

# Test with MCP Inspector
pnpm inspect
```

## License

MIT
