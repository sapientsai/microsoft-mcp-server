import { AuthManager } from "../auth/index.js"
import type { GraphApiRequest, GraphApiResponse, GraphApiVersion } from "../types.js"
import { AZURE_BASE_URL, GRAPH_BASE_URL } from "../types.js"

export class GraphClient {
  private authManager: AuthManager
  private defaultApiVersion: GraphApiVersion

  constructor(authManager: AuthManager, defaultApiVersion: GraphApiVersion = "v1.0") {
    this.authManager = authManager
    this.defaultApiVersion = defaultApiVersion
  }

  async callApi(request: GraphApiRequest): Promise<GraphApiResponse> {
    const accessToken = await this.authManager.getAccessToken()
    if (!accessToken) {
      throw new Error("Not authenticated. Please sign in first.")
    }

    const baseUrl = request.apiType === "azure" ? AZURE_BASE_URL : GRAPH_BASE_URL
    const version = request.apiVersion ?? this.defaultApiVersion
    const method = request.method ?? "GET"

    let url: string
    if (request.apiType === "azure") {
      url = `${baseUrl}${request.path}`
    } else {
      url = `${baseUrl}/${version}${request.path}`
    }

    if (request.queryParams && Object.keys(request.queryParams).length > 0) {
      const params = new URLSearchParams(request.queryParams)
      url += `?${params.toString()}`
    }

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    }

    const fetchOptions: RequestInit = {
      method,
      headers,
    }

    if (request.body && ["POST", "PUT", "PATCH"].includes(method)) {
      fetchOptions.body = JSON.stringify(request.body)
    }

    const response = await fetch(url, fetchOptions)

    const responseHeaders: Record<string, string> = {}
    response.headers.forEach((value, key) => {
      responseHeaders[key] = value
    })

    let data: unknown
    const contentType = response.headers.get("content-type")
    if (contentType?.includes("application/json")) {
      data = await response.json()
    } else {
      data = await response.text()
    }

    if (!response.ok) {
      const errorMessage =
        typeof data === "object" && data !== null && "error" in data
          ? (data as { error: { message?: string } }).error.message
          : `HTTP ${response.status}: ${response.statusText}`
      throw new Error(errorMessage ?? `Request failed with status ${response.status}`)
    }

    return {
      status: response.status,
      data,
      headers: responseHeaders,
    }
  }

  async get(path: string, queryParams?: Record<string, string>): Promise<GraphApiResponse> {
    return this.callApi({ path, method: "GET", queryParams })
  }

  async post(path: string, body?: unknown, queryParams?: Record<string, string>): Promise<GraphApiResponse> {
    return this.callApi({ path, method: "POST", body, queryParams })
  }

  async put(path: string, body?: unknown, queryParams?: Record<string, string>): Promise<GraphApiResponse> {
    return this.callApi({ path, method: "PUT", body, queryParams })
  }

  async patch(path: string, body?: unknown, queryParams?: Record<string, string>): Promise<GraphApiResponse> {
    return this.callApi({ path, method: "PATCH", body, queryParams })
  }

  async delete(path: string, queryParams?: Record<string, string>): Promise<GraphApiResponse> {
    return this.callApi({ path, method: "DELETE", queryParams })
  }
}
