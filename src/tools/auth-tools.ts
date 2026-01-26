import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"

import { AuthManager } from "../auth/index.js"

export function registerAuthTools(server: McpServer, authManager: AuthManager): void {
  server.tool(
    "get_auth_status",
    "Check the current authentication status including mode, token validity, and scopes.",
    {},
    async () => {
      try {
        const status = await authManager.getAuthStatus()
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(status, null, 2),
            },
          ],
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error occurred"
        return {
          content: [
            {
              type: "text" as const,
              text: `Error checking auth status: ${message}`,
            },
          ],
          isError: true,
        }
      }
    },
  )

  server.tool(
    "set_access_token",
    "Manually set an access token for authentication. Use this when you have a pre-obtained bearer token.",
    {
      accessToken: z.string().describe("The bearer access token"),
      expiresOn: z.string().optional().describe("ISO datetime when the token expires (defaults to 1 hour from now)"),
    },
    async ({ accessToken, expiresOn }) => {
      try {
        const expirationDate = expiresOn ? new Date(expiresOn) : undefined
        authManager.setAccessToken(accessToken, expirationDate)

        return {
          content: [
            {
              type: "text" as const,
              text: "Access token set successfully. Auth mode switched to client_token.",
            },
          ],
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error occurred"
        return {
          content: [
            {
              type: "text" as const,
              text: `Error setting access token: ${message}`,
            },
          ],
          isError: true,
        }
      }
    },
  )

  server.tool(
    "sign_in",
    "Initiate device code authentication flow. Returns a code and URL for the user to complete sign-in.",
    {},
    async () => {
      try {
        let deviceCodeMessage = ""

        const tokenInfo = await authManager.initiateDeviceCodeFlow((info) => {
          deviceCodeMessage = `To sign in, visit: ${info.verificationUri}\nEnter code: ${info.userCode}\n\n${info.message}`
        })

        return {
          content: [
            {
              type: "text" as const,
              text: deviceCodeMessage
                ? `${deviceCodeMessage}\n\nAuthentication successful! Signed in as: ${tokenInfo.account ?? "unknown"}`
                : `Authentication successful! Signed in as: ${tokenInfo.account ?? "unknown"}`,
            },
          ],
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error occurred"
        return {
          content: [
            {
              type: "text" as const,
              text: `Sign in failed: ${message}`,
            },
          ],
          isError: true,
        }
      }
    },
  )

  server.tool("sign_out", "Sign out and clear all cached authentication tokens.", {}, async () => {
    try {
      await authManager.signOut()

      return {
        content: [
          {
            type: "text" as const,
            text: "Signed out successfully. All cached tokens have been cleared.",
          },
        ],
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error occurred"
      return {
        content: [
          {
            type: "text" as const,
            text: `Sign out failed: ${message}`,
          },
        ],
        isError: true,
      }
    }
  })
}
