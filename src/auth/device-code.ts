import { AuthenticationResult, Configuration, DeviceCodeRequest, PublicClientApplication } from "@azure/msal-node"

import type { DeviceCodeInfo, TokenInfo } from "../types.js"

export class DeviceCodeAuth {
  private pca: PublicClientApplication
  private cachedToken: TokenInfo | null = null
  private scopes: string[]

  constructor(clientId: string, tenantId: string, scopes: string[]) {
    const msalConfig: Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    }
    this.pca = new PublicClientApplication(msalConfig)
    this.scopes = scopes.map((s) => (s.includes("://") ? s : `https://graph.microsoft.com/${s}`))
  }

  async getAccessToken(): Promise<TokenInfo | null> {
    if (this.cachedToken && this.cachedToken.expiresOn > new Date()) {
      return this.cachedToken
    }

    const accounts = await this.pca.getTokenCache().getAllAccounts()
    if (accounts.length > 0) {
      try {
        const silentResult = await this.pca.acquireTokenSilent({
          account: accounts[0],
          scopes: this.scopes,
        })
        if (silentResult) {
          this.cachedToken = this.toTokenInfo(silentResult)
          return this.cachedToken
        }
      } catch {
        // Silent token acquisition failed, need interactive auth
      }
    }

    return null
  }

  async initiateDeviceCodeFlow(onDeviceCode: (info: DeviceCodeInfo) => void): Promise<TokenInfo> {
    const deviceCodeRequest: DeviceCodeRequest = {
      scopes: this.scopes,
      deviceCodeCallback: (response) => {
        onDeviceCode({
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          message: response.message,
          expiresOn: new Date(Date.now() + response.expiresIn * 1000),
        })
      },
    }

    const result = await this.pca.acquireTokenByDeviceCode(deviceCodeRequest)
    if (!result) {
      throw new Error("Device code authentication failed")
    }

    this.cachedToken = this.toTokenInfo(result)
    return this.cachedToken
  }

  async signOut(): Promise<void> {
    const accounts = await this.pca.getTokenCache().getAllAccounts()
    for (const account of accounts) {
      await this.pca.getTokenCache().removeAccount(account)
    }
    this.cachedToken = null
  }

  private toTokenInfo(result: AuthenticationResult): TokenInfo {
    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn ?? new Date(Date.now() + 3600 * 1000),
      scopes: result.scopes,
      account: result.account?.username,
    }
  }
}
