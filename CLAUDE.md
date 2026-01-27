# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Microsoft Graph MCP Server - A Model Context Protocol server built with FastMCP that provides AI assistants access to Microsoft Graph API for Microsoft 365 services (users, mail, calendar, files, etc.).

## Development Commands

```bash
pnpm validate        # Main command: format + lint + test + build (use before commits)

pnpm test            # Run tests once
pnpm test -- --testNamePattern="pattern"    # Filter by test name
pnpm test -- test/specific.spec.ts          # Run specific file

pnpm build           # Production build
pnpm dev             # Development build with watch mode
pnpm start           # Run the server (requires build first)
```

## Architecture

Built with [FastMCP](https://github.com/punkpeye/fastmcp) - a TypeScript framework for MCP servers.

```
src/
├── index.ts          # Server factory, config, tools, exports
└── bin.ts            # CLI entry point (shebang, dotenv, runServer)
```

### Key Components

- **FastMCP**: Server framework handling MCP protocol, sessions, transport
- **AzureProvider**: OAuth 2.0 authentication with Azure AD/Entra ID
- **Tools**: `microsoft_graph` (Graph API calls) and `get_auth_status` (check auth)

### Authentication Flow

OAuth 2.0 handled by FastMCP's AzureProvider:

1. MCP client (Claude Desktop) initiates OAuth when tool requires auth
2. User redirected to Microsoft login
3. Callback to `{BASE_URL}/oauth/callback`
4. Token stored in session, passed to tools via `session.accessToken`

### Transport Modes

- **httpStream** (default): HTTP server with `/mcp` endpoint and `/oauth/callback`
- **stdio**: For Claude Code CLI integration

## Configuration

Required environment variables:

```bash
AZURE_CLIENT_ID=...      # Azure app registration client ID
AZURE_CLIENT_SECRET=...  # Azure app client secret
AZURE_TENANT_ID=common   # Tenant ID (default: "common" for multi-tenant)
BASE_URL=http://localhost:8080  # Server URL (for OAuth callback)
PORT=8080                # Server port
TRANSPORT_TYPE=httpStream  # httpStream or stdio
```

## Build System

Uses `ts-builds` toolchain with `tsdown` bundler. Configuration extends shared configs:

- `tsconfig.json` extends `ts-builds/tsconfig`
- `tsdown.config.ts` imports from `ts-builds/tsdown`
- Prettier uses `ts-builds/prettier`

Output: ES modules in `dist/` with TypeScript declarations.
