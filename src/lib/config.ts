export const APP_VERSION = "1.0.0";

export const MICROSOFT_ADS_SCOPE = "https://ads.microsoft.com/msads.manage";

export const MICROSOFT_OAUTH_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  MICROSOFT_ADS_SCOPE,
] as const;

export const MICROSOFT_AUTHORIZE_URL =
  "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

export const MICROSOFT_TOKEN_URL =
  "https://login.microsoftonline.com/common/oauth2/v2.0/token";

export const DEFAULT_FIRESTORE_DATABASE_ID = "bing-mcp-v1";

export const FORBIDDEN_FIRESTORE_DATABASE_IDS = new Set([
  "(default)",
  "default",
  "gconnect-mcp-all-v1",
  "ga4-mcp-v3",
  "ga4-mcp-v2",
  "ga4-mcp-v1",
]);

export interface AppConfig {
  appBaseUrl: string;
  microsoftClientId: string;
  microsoftClientSecret: string;
  microsoftRedirectUri: string;
  microsoftAdsDeveloperToken: string;
  mcpTokenSecret: string;
  oauthStateSecret: string;
  tokenEncryptionKey: string;
  firestoreProjectId: string | undefined;
  firestoreDatabaseId: string;
  mcpOAuthClientId: string | undefined;
  mcpOAuthClientSecret: string | undefined;
}

function required(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function stripTrailingSlash(url: string): string {
  return url.replace(/\/+$/, "");
}

export function resolveFirestoreDatabaseId(raw?: string): string {
  const value = (raw ?? "").trim();
  if (!value) {
    throw new Error("Missing required environment variable: FIRESTORE_DATABASE_ID");
  }
  const normalized = value.toLowerCase();
  if (FORBIDDEN_FIRESTORE_DATABASE_IDS.has(value) || FORBIDDEN_FIRESTORE_DATABASE_IDS.has(normalized)) {
    throw new Error(
      `FIRESTORE_DATABASE_ID must be the named database "${DEFAULT_FIRESTORE_DATABASE_ID}". ` +
        `Do not use "(default)", "gconnect-mcp-all-v1", or any other product database.`,
    );
  }
  if (value !== DEFAULT_FIRESTORE_DATABASE_ID) {
    throw new Error(
      `FIRESTORE_DATABASE_ID must be "${DEFAULT_FIRESTORE_DATABASE_ID}" so this app does not conflict with other Firestore databases.`,
    );
  }
  return value;
}

export function getConfig(): AppConfig {
  const appBaseUrl = stripTrailingSlash(required("APP_BASE_URL"));
  const microsoftRedirectUri =
    process.env.MICROSOFT_REDIRECT_URI?.trim() || `${appBaseUrl}/oauth/microsoft/callback`;

  return {
    appBaseUrl,
    microsoftClientId: required("MICROSOFT_CLIENT_ID"),
    microsoftClientSecret: required("MICROSOFT_CLIENT_SECRET"),
    microsoftRedirectUri,
    microsoftAdsDeveloperToken: required("MICROSOFT_ADS_DEVELOPER_TOKEN"),
    mcpTokenSecret: required("MCP_TOKEN_SECRET"),
    oauthStateSecret: required("OAUTH_STATE_SECRET"),
    tokenEncryptionKey: required("TOKEN_ENCRYPTION_KEY"),
    firestoreProjectId: process.env.FIRESTORE_PROJECT_ID?.trim() || undefined,
    firestoreDatabaseId: resolveFirestoreDatabaseId(process.env.FIRESTORE_DATABASE_ID),
    mcpOAuthClientId: process.env.MCP_OAUTH_CLIENT_ID?.trim() || undefined,
    mcpOAuthClientSecret: process.env.MCP_OAUTH_CLIENT_SECRET?.trim() || undefined,
  };
}
