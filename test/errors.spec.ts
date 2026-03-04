import { describe, expect, it } from "vitest"

import { type AppError, authError, configError, graphError } from "../src/errors.js"

describe("Error factories", () => {
  describe("graphError", () => {
    it("should create a GraphError with correct fields", () => {
      const err = graphError("NotFound", "Resource not found", 404)

      expect(err).toEqual({
        _kind: "GraphError",
        code: "NotFound",
        message: "Resource not found",
        status: 404,
      })
    })

    it("should be readonly", () => {
      const err = graphError("Forbidden", "Access denied", 403)

      expect(err._kind).toBe("GraphError")
      expect(err.code).toBe("Forbidden")
      expect(err.message).toBe("Access denied")
      expect(err.status).toBe(403)
    })
  })

  describe("authError", () => {
    it("should create an AuthError with correct fields", () => {
      const err = authError("Not authenticated")

      expect(err).toEqual({
        _kind: "AuthError",
        message: "Not authenticated",
      })
    })
  })

  describe("configError", () => {
    it("should create a ConfigError with correct fields", () => {
      const err = configError("Missing tenant ID")

      expect(err).toEqual({
        _kind: "ConfigError",
        message: "Missing tenant ID",
      })
    })
  })

  describe("AppError discrimination", () => {
    it("should discriminate error types by _kind", () => {
      const errors: readonly AppError[] = [
        graphError("BadRequest", "Invalid query", 400),
        authError("Token expired"),
        configError("Invalid config"),
      ]

      const graph = errors.filter((e) => e._kind === "GraphError")
      const auth = errors.filter((e) => e._kind === "AuthError")
      const config = errors.filter((e) => e._kind === "ConfigError")

      expect(graph).toHaveLength(1)
      expect(auth).toHaveLength(1)
      expect(config).toHaveLength(1)

      expect(graph[0]?._kind).toBe("GraphError")
      expect(auth[0]?._kind).toBe("AuthError")
      expect(config[0]?._kind).toBe("ConfigError")
    })

    it("should narrow types with _kind discriminant", () => {
      const err: AppError = graphError("Throttled", "Too many requests", 429)

      if (err._kind === "GraphError") {
        expect(err.code).toBe("Throttled")
        expect(err.status).toBe(429)
      } else {
        throw new Error("Expected GraphError")
      }
    })
  })
})
