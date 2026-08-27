import { createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import {
  getConfig,
  MICROSOFT_AUTHORIZE_URL,
  MICROSOFT_OAUTH_SCOPES,
  MICROSOFT_TOKEN_URL,
} from "@/lib/config";
import {
  AppError,
  AuthenticationError,
  ConnectionReauthorizationRequiredError,
  mapMicrosoftError,
} from "@/lib/errors";
import { decodeJwtPayload } from "@/mcp/oauth/jwt";

const STATE_TTL_MS = 15 * 60 * 1000;

export interface PkcePair {
  verifier: string;
  challenge: string;
}

export interface SignedOAuthState {
  nonce: string;
  operation: string;
  exp: number;
}

export interface MicrosoftIdentity {
  microsoftSubjectId: string;
  email: string | null;
  displayName: string | null;
  refreshToken: string | null;
  accessToken: string | null;
  accessTokenExpiresAt: number | null;
  scopes: string[];
}

interface MicrosoftTokenResponse {
  token_type?: string;
  scope?: string;
  expires_in?: number;
  access_token?: string;
  refresh_token?: string;
  id_token?: string;
  error?: string;
  error_description?: string;
}

interface MicrosoftIdToken {
  sub?: string;
  oid?: string;
  email?: string;
  preferred_username?: string;
  name?: string;
  aud?: string;
}

export function createPkcePair(): PkcePair {
  const verifier = randomBytes(32).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  return { verifier, challenge };
}

export function signOAuthState(
  state: SignedOAuthState,
  secret = getConfig().oauthStateSecret,
): string {
  const payload = Buffer.from(JSON.stringify(state)).toString("base64url");
  const signature = createHmac("sha256", secret).update(payload).digest("base64url");
  return `${payload}.${signature}`;
}

export function verifyOAuthState(
  value: string,
  secret = getConfig().oauthStateSecret,
): SignedOAuthState {
  const [payload, signature] = value.split(".");
  if (!payload || !signature) {
    throw new AuthenticationError("OAuth state is invalid.");
  }

  const expected = createHmac("sha256", secret).update(payload).digest("base64url");
  const signatureBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expected);
  if (
    signatureBuffer.length !== expectedBuffer.length ||
    !timingSafeEqual(signatureBuffer, expectedBuffer)
  ) {
    throw new AuthenticationError("OAuth state is invalid.");
  }

  const parsed = JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as SignedOAuthState;
  if (!parsed.nonce || typeof parsed.exp !== "number" || parsed.exp < Date.now()) {
    throw new AuthenticationError(
      "OAuth state has expired. Start the Microsoft authorization flow again.",
    );
  }
  return parsed;
}

export function createSignedOAuthState(operation: string): string {
  return signOAuthState({
    nonce: randomBytes(16).toString("hex"),
    operation,
    exp: Date.now() + STATE_TTL_MS,
  });
}

export function microsoftScopes(): string[] {
  return [...MICROSOFT_OAUTH_SCOPES];
}

export function parseGrantedScopes(scope?: string | null): string[] {
  const granted = (scope ?? "")
    .split(/\s+/)
    .map((item) => item.trim())
    .filter(Boolean);
  return granted.length > 0 ? [...new Set(granted)] : microsoftScopes();
}

export function buildMicrosoftAuthUrl(params: {
  state: string;
  codeChallenge: string;
  prompt?: "select_account" | "login" | "consent";
}): string {
  const config = getConfig();
  const url = new URL(MICROSOFT_AUTHORIZE_URL);
  url.searchParams.set("client_id", config.microsoftClientId);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", config.microsoftRedirectUri);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", microsoftScopes().join(" "));
  url.searchParams.set("state", params.state);
  url.searchParams.set("code_challenge", params.codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("prompt", params.prompt ?? "select_account");
  return url.toString();
}

function identityFromTokens(tokens: MicrosoftTokenResponse): MicrosoftIdentity {
  if (!tokens.id_token) {
    throw new AppError(
      "Microsoft did not return an ID token. OpenID is required to identify the Microsoft account.",
      "microsoft_oauth_error",
      400,
    );
  }
  const claims = decodeJwtPayload<MicrosoftIdToken>(tokens.id_token);
  const microsoftSubjectId = claims.sub?.trim();
  if (!microsoftSubjectId) {
    throw new AppError("Microsoft identity is missing a subject identifier.", "unauthorized", 400);
  }
  return {
    microsoftSubjectId,
    email: claims.email ?? claims.preferred_username ?? null,
    displayName: claims.name ?? null,
    refreshToken: tokens.refresh_token ?? null,
    accessToken: tokens.access_token ?? null,
    accessTokenExpiresAt: tokens.expires_in
      ? Date.now() + tokens.expires_in * 1000
      : null,
    scopes: parseGrantedScopes(tokens.scope),
  };
}

async function postToken(body: URLSearchParams): Promise<MicrosoftTokenResponse> {
  const response = await fetch(MICROSOFT_TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Accept: "application/json",
    },
    body,
  });
  const tokens = (await response.json()) as MicrosoftTokenResponse;
  if (!response.ok || tokens.error) {
    throw mapMicrosoftError({
      status: response.status,
      error: tokens.error,
      error_description: tokens.error_description,
      message: tokens.error_description ?? tokens.error,
    });
  }
  return tokens;
}

export async function exchangeAuthorizationCode(params: {
  code: string;
  codeVerifier: string;
}): Promise<MicrosoftIdentity> {
  const config = getConfig();
  const tokens = await postToken(
    new URLSearchParams({
      client_id: config.microsoftClientId,
      client_secret: config.microsoftClientSecret,
      grant_type: "authorization_code",
      code: params.code,
      redirect_uri: config.microsoftRedirectUri,
      code_verifier: params.codeVerifier,
      scope: microsoftScopes().join(" "),
    }),
  );
  return identityFromTokens(tokens);
}

export async function refreshMicrosoftAccessToken(refreshToken: string): Promise<{
  accessToken: string;
  expiresAt: number;
  refreshToken: string;
  scopes: string[];
}> {
  const config = getConfig();
  try {
    const tokens = await postToken(
      new URLSearchParams({
        client_id: config.microsoftClientId,
        client_secret: config.microsoftClientSecret,
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope: microsoftScopes().join(" "),
      }),
    );
    if (!tokens.access_token) {
      throw new ConnectionReauthorizationRequiredError(
        "Microsoft did not return an access token. Reauthorize this connection.",
      );
    }
    return {
      accessToken: tokens.access_token,
      expiresAt: Date.now() + (tokens.expires_in ?? 3600) * 1000,
      refreshToken: tokens.refresh_token ?? refreshToken,
      scopes: parseGrantedScopes(tokens.scope),
    };
  } catch (error) {
    throw mapMicrosoftError(error);
  }
}
