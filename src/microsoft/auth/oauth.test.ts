import { beforeEach, describe, expect, it, vi } from "vitest";
import { mapMicrosoftError } from "@/lib/errors";
import {
  createPkcePair,
  createSignedOAuthState,
  exchangeAuthorizationCode,
  refreshMicrosoftAccessToken,
  verifyOAuthState,
} from "@/microsoft/auth/oauth";
import { setRequiredEnv } from "@/test/env";

describe("Microsoft OAuth", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("signs and verifies OAuth state", () => {
    const state = createSignedOAuthState("bootstrap_operator");
    const parsed = verifyOAuthState(state);
    expect(parsed.operation).toBe("bootstrap_operator");
    expect(parsed.nonce).toMatch(/^[0-9a-f]+$/);
  });

  it("creates a PKCE pair", () => {
    const pair = createPkcePair();
    expect(pair.verifier).toBeTruthy();
    expect(pair.challenge).toBeTruthy();
    expect(pair.challenge).not.toBe(pair.verifier);
  });

  it("exchanges an authorization code for Microsoft identity", async () => {
    const idTokenPayload = Buffer.from(
      JSON.stringify({
        sub: "ms-sub-1",
        email: "user@example.com",
        name: "Example User",
        aud: "test-microsoft-client-id",
      }),
    ).toString("base64url");
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          access_token: "access-1",
          refresh_token: "refresh-1",
          expires_in: 3600,
          scope: "openid profile email offline_access https://ads.microsoft.com/msads.manage",
          id_token: `header.${idTokenPayload}.sig`,
        }),
      }),
    );

    const identity = await exchangeAuthorizationCode({
      code: "auth-code",
      codeVerifier: "verifier",
    });
    expect(identity.microsoftSubjectId).toBe("ms-sub-1");
    expect(identity.email).toBe("user@example.com");
    expect(identity.refreshToken).toBe("refresh-1");
  });

  it("refreshes an access token", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          access_token: "new-access",
          refresh_token: "new-refresh",
          expires_in: 3600,
          scope: "https://ads.microsoft.com/msads.manage",
        }),
      }),
    );
    const refreshed = await refreshMicrosoftAccessToken("old-refresh");
    expect(refreshed.accessToken).toBe("new-access");
    expect(refreshed.refreshToken).toBe("new-refresh");
  });

  it("maps invalid_grant to reauthorization required", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: false,
        status: 400,
        json: async () => ({
          error: "invalid_grant",
          error_description: "The refresh token has been revoked.",
        }),
      }),
    );
    await expect(refreshMicrosoftAccessToken("revoked")).rejects.toMatchObject({
      code: "reauthorization_required",
    });
  });

  it("maps invalid_grant text without a fetch", () => {
    const mapped = mapMicrosoftError({
      error: "invalid_grant",
      error_description: "AADSTS70000: The refresh token is invalid",
    });
    expect(mapped.code).toBe("reauthorization_required");
  });
});
