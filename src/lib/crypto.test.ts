import { beforeEach, describe, expect, it } from "vitest";
import { decryptRefreshToken, encryptRefreshToken } from "@/lib/crypto";
import { setRequiredEnv } from "@/test/env";

describe("refresh token encryption", () => {
  beforeEach(() => {
    setRequiredEnv();
  });

  it("round-trips a refresh token", () => {
    const encrypted = encryptRefreshToken("0.secret-refresh-token");
    expect(encrypted).not.toContain("0.secret-refresh-token");
    expect(decryptRefreshToken(encrypted)).toBe("0.secret-refresh-token");
  });
});
