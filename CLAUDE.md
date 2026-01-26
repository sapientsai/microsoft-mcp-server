# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Microsoft Graph MCP Server - A Model Context Protocol server that provides AI assistants access to Microsoft Graph API for Microsoft 365 services (users, mail, calendar, files, etc.).

## Development Commands

```bash
pnpm validate        # Main command: format + lint + test + build (use before commits)

pnpm test            # Run tests once
pnpm test -- --testNamePattern="pattern"    # Filter by test name
pnpm test -- test/specific.spec.ts          # Run specific file

pnpm build           # Production build
pnpm dev             # Development build with watch mode
pnpm start           # Run the server (requires build first)
pnpm inspect         # Test with MCP Inspector
```

## Architecture

### MCP Server Structure

```
src/
├── index.ts          # Server factory, config loader, exports
├── bin.ts            # CLI entry point (shebang, dotenv, runServer)
├── types.ts          # TypeScript types and constants
├── auth/             # Authentication module
│   ├── auth-manager.ts      # Coordinates auth modes, token management
│   ├── device-code.ts       # MSAL device code flow (interactive)
│   └── client-credentials.ts # MSAL client credentials (app-only)
├── client/
│   └── graph-client.ts      # HTTP client wrapper for Graph API
└── tools/            # MCP tool definitions
    ├── graph-tools.ts       # microsoft_graph tool
    └── auth-tools.ts        # get_auth_status, set_access_token, sign_in, sign_out
```

### Authentication Flow

Three modes supported via `AUTH_MODE` environment variable:

- **device_code**: Interactive browser auth, user signs in via microsoft.com/devicelogin
- **client_credentials**: App-only auth using client ID + secret (no user context)
- **client_token**: Manual token injection via env var or `set_access_token` tool

`AuthManager` coordinates mode switching and token lifecycle. Device code and client credentials use `@azure/msal-node` for token acquisition.

### Tool Registration Pattern

Tools are registered via `server.tool()` from `@modelcontextprotocol/sdk`. Each tool:

1. Defines name, description, and Zod schema for parameters
2. Returns `{ content: [{ type: "text", text: "..." }] }` format
3. Sets `isError: true` on failures

### Graph Client

`GraphClient` wraps fetch calls with:

- Bearer token injection from AuthManager
- Base URL construction (Graph vs Azure ARM)
- API version handling (v1.0 or beta)
- OData query parameter serialization

## Configuration

Required environment variables (see `.env.example`):

- `AZURE_CLIENT_ID` - Azure app registration client ID
- `AZURE_TENANT_ID` - Tenant ID (default: "common")
- `AUTH_MODE` - Authentication mode
- `AZURE_CLIENT_SECRET` - Required for client_credentials mode
- `ACCESS_TOKEN` - Pre-obtained token for client_token mode

## Build System

Uses `ts-builds` toolchain with `tsdown` bundler. Configuration extends shared configs:

- `tsconfig.json` extends `ts-builds/tsconfig`
- `tsdown.config.ts` imports from `ts-builds/tsdown`
- Prettier uses `ts-builds/prettier`

Output: ES modules in `dist/` with TypeScript declarations.
