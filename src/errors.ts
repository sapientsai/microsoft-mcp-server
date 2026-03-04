export type GraphError = {
  readonly _kind: "GraphError"
  readonly code: string
  readonly message: string
  readonly status: number
}

export type AuthError = {
  readonly _kind: "AuthError"
  readonly message: string
}

export type ConfigError = {
  readonly _kind: "ConfigError"
  readonly message: string
}

export type AppError = GraphError | AuthError | ConfigError

export const graphError = (code: string, message: string, status: number): GraphError => ({
  _kind: "GraphError",
  code,
  message,
  status,
})

export const authError = (message: string): AuthError => ({
  _kind: "AuthError",
  message,
})

export const configError = (message: string): ConfigError => ({
  _kind: "ConfigError",
  message,
})
