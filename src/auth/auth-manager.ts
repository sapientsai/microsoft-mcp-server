import type { AuthMode, AuthStatus, DeviceCodeInfo, ServerConfig, TokenInfo } from "../types.js"
import { ClientCredentialsAuth } from "./client-credentials.js"
import { DeviceCodeAuth } from "./device-code.js"

export class AuthManager {
  private config: ServerConfig
  private deviceCodeAuth: DeviceCodeAuth | null = null
  private clientCredentialsAuth: ClientCredentialsAuth | null = null
  private manualToken: TokenInfo | null = null
  private currentMode: AuthMode

  constructor(config: ServerConfig) {
    this.config = config
    this.currentMode = config.authMode

    if (config.authMode === "device_code") {
      this.deviceCodeAuth = new DeviceCodeAuth(config.clientId, config.tenantId, config.scopes)
    } else if (config.authMode === "client_credentials" && config.clientSecret) {
      this.clientCredentialsAuth = new ClientCredentialsAuth(
        config.clientId,
        config.tenantId,
        config.clientSecret,
        config.scopes,
      )
    }
  }

  async getAccessToken(): Promise<string | null> {
    const tokenInfo = await this.getTokenInfo()
    return tokenInfo?.accessToken ?? null
  }

  async getTokenInfo(): Promise<TokenInfo | null> {
    switch (this.currentMode) {
      case "client_token":
        if (this.manualToken && this.manualToken.expiresOn > new Date()) {
          return this.manualToken
        }
        return null

      case "device_code":
        if (!this.deviceCodeAuth) {
          this.deviceCodeAuth = new DeviceCodeAuth(this.config.clientId, this.config.tenantId, this.config.scopes)
        }
        return this.deviceCodeAuth.getAccessToken()

      case "client_credentials":
        if (!this.clientCredentialsAuth) {
          if (!this.config.clientSecret) {
            throw new Error("Client secret required for client_credentials mode")
          }
          this.clientCredentialsAuth = new ClientCredentialsAuth(
            this.config.clientId,
            this.config.tenantId,
            this.config.clientSecret,
            this.config.scopes,
          )
        }
        return this.clientCredentialsAuth.getAccessToken()
    }
  }

  async getAuthStatus(): Promise<AuthStatus> {
    const tokenInfo = await this.getTokenInfo()

    return {
      authenticated: tokenInfo !== null,
      mode: this.currentMode,
      expiresOn: tokenInfo?.expiresOn.toISOString(),
      scopes: tokenInfo?.scopes,
      account: tokenInfo?.account,
    }
  }

  setAccessToken(accessToken: string, expiresOn?: Date): void {
    this.currentMode = "client_token"
    this.manualToken = {
      accessToken,
      expiresOn: expiresOn ?? new Date(Date.now() + 3600 * 1000),
    }
  }

  async initiateDeviceCodeFlow(onDeviceCode: (info: DeviceCodeInfo) => void): Promise<TokenInfo> {
    this.currentMode = "device_code"
    if (!this.deviceCodeAuth) {
      this.deviceCodeAuth = new DeviceCodeAuth(this.config.clientId, this.config.tenantId, this.config.scopes)
    }
    return this.deviceCodeAuth.initiateDeviceCodeFlow(onDeviceCode)
  }

  async signOut(): Promise<void> {
    if (this.deviceCodeAuth) {
      await this.deviceCodeAuth.signOut()
    }
    if (this.clientCredentialsAuth) {
      this.clientCredentialsAuth.clearCache()
    }
    this.manualToken = null
  }

  getMode(): AuthMode {
    return this.currentMode
  }

  setMode(mode: AuthMode): void {
    this.currentMode = mode
  }
}
