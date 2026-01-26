import { ClientCredentialRequest, ConfidentialClientApplication, Configuration } from "@azure/msal-node"

import type { TokenInfo } from "../types.js"

export class ClientCredentialsAuth {
  private cca: ConfidentialClientApplication
  private cachedToken: TokenInfo | null = null
  private scopes: string[]

  constructor(clientId: string, tenantId: string, clientSecret: string, scopes: string[]) {
    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret,
      },
    }
    this.cca = new ConfidentialClientApplication(msalConfig)
    // Client credentials use .default scope for app-only permissions
    this.scopes = scopes.length > 0 ? scopes : ["https://graph.microsoft.com/.default"]
  }

  async getAccessToken(): Promise<TokenInfo | null> {
    if (this.cachedToken && this.cachedToken.expiresOn > new Date()) {
      return this.cachedToken
    }

    try {
      const request: ClientCredentialRequest = {
        scopes: this.scopes,
      }

      const result = await this.cca.acquireTokenByClientCredential(request)
      if (!result) {
        return null
      }

      this.cachedToken = {
        accessToken: result.accessToken,
        expiresOn: result.expiresOn ?? new Date(Date.now() + 3600 * 1000),
        scopes: result.scopes,
      }

      return this.cachedToken
    } catch (error) {
      console.error("Client credentials authentication failed:", error)
      return null
    }
  }

  clearCache(): void {
    this.cachedToken = null
  }
}
