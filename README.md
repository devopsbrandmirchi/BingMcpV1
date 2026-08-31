# Microsoft Advertising MCP Connector

Standalone **True Model B** MCP server for Claude.ai Custom Connectors. One operator can connect multiple Microsoft accounts. Each connection can see multiple Microsoft Advertising customers and advertising accounts.

This is **not** part of the Google All-in-One connector. It has no runtime dependency on that project, no shared Firestore collections, and no shared OAuth tokens.

```text
Claude.ai
   │  MCP OAuth 2.1
   ▼
operatorId
   ├── Microsoft Connection A → customers → advertising accounts → campaigns / reports
   ├── Microsoft Connection B
   └── Microsoft Connection C
```

MCP JWT `sub` is always the internal `operatorId`. Microsoft OpenID `sub` lives only on `bing_mcp_microsoft_connections.microsoftSubjectId`. Email is display metadata, never an authorization key.

## Architecture

- Next.js 15 App Router + `mcp-handler` Streamable HTTP at `/mcp`
- MCP OAuth 2.1 with PKCE S256, DCR, and Claude CIMD
- Microsoft Entra authorization-code flow (`common` tenant)
- Microsoft Advertising **REST v13** (not SOAP)
- Named Firestore database `bing-mcp-v1`
- AES-256-GCM refresh-token encryption at rest
- Cloud Run (stateless; listen on `$PORT`)

## Credentials (these are different things)

| Variable | What it is |
|---|---|
| `MICROSOFT_CLIENT_ID` / `MICROSOFT_CLIENT_SECRET` | Entra app registration. Application-level. Used to run the Microsoft OAuth flow. |
| `MICROSOFT_ADS_DEVELOPER_TOKEN` | Microsoft Advertising developer token. Application-level. Required on every Ads API call. **Not** per operator. Never returned to Claude. |
| Per-connection refresh token | User OAuth credential stored encrypted on that operator’s Microsoft connection. Connecting account Y does not replace account X. |

## Environment

Copy `.env.example` to `.env`. Never commit `.env`.

Required:

```text
APP_BASE_URL
MICROSOFT_CLIENT_ID
MICROSOFT_CLIENT_SECRET
MICROSOFT_REDIRECT_URI
MICROSOFT_ADS_DEVELOPER_TOKEN
MCP_TOKEN_SECRET
OAUTH_STATE_SECRET
TOKEN_ENCRYPTION_KEY
FIRESTORE_PROJECT_ID
FIRESTORE_DATABASE_ID=bing-mcp-v1
```

`FIRESTORE_DATABASE_ID` must be `bing-mcp-v1`. Startup rejects `(default)` and `gconnect-mcp-all-v1`.

## Microsoft Entra app registration

1. Open [Microsoft Entra app registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).
2. Create a registration.
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts** (so `common` works).
4. Add a Web redirect URI: `{APP_BASE_URL}/oauth/microsoft/callback` (local example: `http://localhost:3000/oauth/microsoft/callback`).
5. Create a client secret. Store it in `MICROSOFT_CLIENT_SECRET`.
6. API permissions: Microsoft Advertising delegated scope `https://ads.microsoft.com/msads.manage`. Also request `openid`, `profile`, `email`, and `offline_access`.
7. Grant admin consent if your tenant requires it.

Official guides:

- [Authentication with OAuth](https://learn.microsoft.com/en-us/advertising/guides/authentication-oauth?view=bingads-13)
- [Request user consent](https://learn.microsoft.com/en-us/advertising/guides/authentication-oauth-consent?view=bingads-13)

## Developer token

Get a Microsoft Advertising developer token from Microsoft Advertising Developer Settings. Sandbox tokens can be self-issued; production tokens require Microsoft approval.

This token is an **application** credential. It is sent as the `DeveloperToken` header on Advertising REST calls. It is never stored per operator and must never appear in MCP output or logs.

## Local development

```powershell
npm install
npm test
npm run typecheck
npm run dev
```

- Health: `GET http://localhost:3000/health`
- MCP: `http://localhost:3000/mcp`
- Microsoft callback: `http://localhost:3000/oauth/microsoft/callback`

Add the Custom Connector in Claude.ai pointing at `{APP_BASE_URL}/mcp`. Claude completes MCP OAuth, then the server redirects the browser to Microsoft login. After consent, continue back to Claude.

Firestore locally: either Application Default Credentials (`gcloud auth application-default login`) plus `FIRESTORE_PROJECT_ID`, or `FIRESTORE_EMULATOR_HOST`.

## MCP tools (V1, read-only)

Connections: `get_operator`, `list_microsoft_connections`, `get_microsoft_connection`, `start_microsoft_connection`, `disconnect_microsoft_connection`

Customers / accounts: `list_microsoft_customers`, `get_microsoft_customer`, `list_microsoft_accounts`, `get_microsoft_account`

Entities: `list_campaigns`, `get_campaign`, `list_ad_groups`, `get_ad_group`, `list_ads`, `get_ad`, `list_keywords`, `get_keyword`, `list_conversion_goals`, `get_conversion_goal`, `list_uet_tags`

Reports (official Reporting API, async submit/poll/download): `get_account_performance`, `get_campaign_performance`, `get_ad_group_performance`, `get_keyword_performance`

If the same advertising account is visible through more than one Microsoft connection, the server returns an ambiguity error and asks for `connectionId`. It never guesses.

## Firestore collections

Named database: `bing-mcp-v1`

- `bing_mcp_operators`
- `bing_mcp_microsoft_connections`
- `bing_mcp_connection_uniques` (`{operatorId}_{microsoftSubjectId}`)
- `bing_mcp_microsoft_customers`
- `bing_mcp_microsoft_accounts`
- `bing_mcp_oauth_transactions`
- `bing_mcp_sessions`

## Cloud Run

Cloud Run ingress must be **unauthenticated**. Authentication is MCP JWT, not Cloud IAM.

The Cloud Run service account needs Firestore access on the named database (`roles/datastore.user` is typical). Use Application Default Credentials. Do not commit a service-account JSON file.

```powershell
.\scripts\cloud-run-setup.ps1 -ProjectId YOUR_PROJECT
.\scripts\cloud-run-deploy.ps1 -ProjectId YOUR_PROJECT
.\scripts\cloud-run-set-env.ps1 -ProjectId YOUR_PROJECT -AppBaseUrl https://... -MicrosoftClientId ... -MicrosoftClientSecret ... -McpTokenSecret ... -OauthStateSecret ... -TokenEncryptionKey ... -MicrosoftAdsDeveloperToken ...
```

Then add `{APP_BASE_URL}/oauth/microsoft/callback` to the Entra app and add the Claude Custom Connector at `{APP_BASE_URL}/mcp`.

Service timeout is 300 seconds so async reports can finish.

## Security

Never exposed through MCP or logs: access tokens, refresh tokens, client secrets, developer token, encryption keys, authorization codes.

Authorization uses `operatorId` + `connectionId` + Microsoft IDs. Names and emails are not authorization identifiers.

## V2 (not implemented)

Campaign create/update/pause, ad group and ad writes, keyword writes, budget updates.

## Tests

```powershell
npm test
```

Automated tests cover MCP JWT identity, Microsoft OAuth state/exchange/refresh/`invalid_grant`, operator isolation, multi-connection, same Microsoft identity across operators, ambiguous account rejection, report CSV parsing, header rules, and secret redaction.

Live Claude, Microsoft login, and Cloud Run paths require your Entra app, developer token, and GCP project.
