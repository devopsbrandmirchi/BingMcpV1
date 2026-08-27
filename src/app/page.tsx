export const dynamic = "force-dynamic";

export default function HomePage() {
  const baseUrl = (process.env.APP_BASE_URL ?? "").replace(/\/+$/, "");
  const mcpUrl = baseUrl ? `${baseUrl}/mcp` : "/mcp";
  const microsoftCallback = baseUrl
    ? `${baseUrl}/oauth/microsoft/callback`
    : "/oauth/microsoft/callback";

  return (
    <main
      style={{
        fontFamily: "ui-sans-serif, system-ui, sans-serif",
        maxWidth: "42rem",
        margin: "3rem auto",
        padding: "0 1.25rem",
        lineHeight: 1.5,
      }}
    >
      <h1>Microsoft Advertising MCP Connector</h1>
      <p>
        Standalone True Model B connector: one MCP operator can own multiple
        independent Microsoft OAuth connections. Advertising customers and
        accounts are routed to the Microsoft credential that authorized them.
      </p>
      <ul>
        <li>
          Claude.ai connector URL: <code>{mcpUrl}</code>
        </li>
        <li>
          Microsoft OAuth callback: <code>{microsoftCallback}</code>
        </li>
      </ul>
      <p>
        <a href="/health">Health</a>
      </p>
      <p style={{ color: "#666" }}>
        Claude authenticates with MCP OAuth. The MCP token subject is the
        internal operator ID, never a Microsoft account. This service uses the
        named Firestore database <code>bing-mcp-v1</code> and must allow
        unauthenticated Cloud Run ingress so Anthropic can reach <code>/mcp</code>.
      </p>
    </main>
  );
}
