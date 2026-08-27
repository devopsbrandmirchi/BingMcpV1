import { beforeEach, describe, expect, it } from "vitest";
import { tryReadAccessToken } from "@/mcp/auth";
import { issueAccessToken } from "@/mcp/oauth/tokens";
import { unauthorizedMcpResponse } from "@/mcp/server";
import { setRequiredEnv } from "@/test/env";

describe("MCP authentication", () => {
  beforeEach(() => {
    setRequiredEnv();
  });

  it("rejects missing tokens", () => {
    expect(tryReadAccessToken(undefined)).toBeUndefined();
  });

  it("accepts a valid operator access token", () => {
    const token = issueAccessToken({ clientId: "claude", sub: "operator-1" });
    expect(tryReadAccessToken(token)?.sub).toBe("operator-1");
  });

  it("returns 401 without leaking secrets", async () => {
    const response = unauthorizedMcpResponse();
    const body = await response.json();
    expect(response.status).toBe(401);
    expect(JSON.stringify(body)).not.toMatch(/developer|refresh|secret/i);
  });
});
