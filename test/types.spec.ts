import { describe, expect, it } from "vitest"
import { DEFAULT_SCOPES, GRAPH_BASE_URL, AZURE_BASE_URL } from "../src/types.js"

describe("Types", () => {
  describe("DEFAULT_SCOPES", () => {
    it("should contain essential scopes", () => {
      expect(DEFAULT_SCOPES).toContain("User.Read")
      expect(DEFAULT_SCOPES).toContain("Mail.Read")
      expect(DEFAULT_SCOPES).toContain("Calendars.Read")
    })

    it("should be an array", () => {
      expect(Array.isArray(DEFAULT_SCOPES)).toBe(true)
    })
  })

  describe("Constants", () => {
    it("should have correct Graph base URL", () => {
      expect(GRAPH_BASE_URL).toBe("https://graph.microsoft.com")
    })

    it("should have correct Azure base URL", () => {
      expect(AZURE_BASE_URL).toBe("https://management.azure.com")
    })
  })
})
