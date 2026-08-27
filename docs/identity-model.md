# Identity model

True Model B for Microsoft Advertising.

## Two identities

| Identity | Field | Source |
|----------|-------|--------|
| MCP operator | `operatorId` | MCP JWT `sub` |
| Microsoft account | `microsoftSubjectId` | Microsoft OpenID `sub` on `bing_mcp_microsoft_connections` |

These are never interchangeable.

## Multi-connection

One operator can connect Microsoft accounts X, Y, and Z independently.

The same Microsoft identity connected by Operator A and Operator B produces two connection documents. They do not share refresh tokens.

Uniqueness key: `bing_mcp_connection_uniques/{operatorId}_{microsoftSubjectId}`.

Email is display metadata only.

## Resume vs start new

During first MCP authorize, if the signed-in Microsoft subject already has connections on other operators, the user chooses **Resume** or **Start new**. Connections are never merged.
