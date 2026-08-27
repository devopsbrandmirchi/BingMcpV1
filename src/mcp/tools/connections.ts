import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import { connectUrl, issueConnectToken } from "@/microsoft/auth/connect-token";
import { jsonToolResult, connectionIdSchema } from "@/mcp/tools/schemas";
import { loadCurrentOperatorRecord, resolveOwnedConnection } from "@/services/resolver";
import { publicConnectionView } from "@/services/tokenService";
import { getAppStore } from "@/store/app-store";

export function registerConnectionTools(server: McpServer): void {
  server.registerTool(
    "list_microsoft_connections",
    {
      title: "List Microsoft connections",
      description:
        "Lists Microsoft OAuth connections owned by the authenticated operator. Does not return tokens or secrets.",
      inputSchema: z.object({}),
    },
    async () => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const connections = await getAppStore().listConnections(operator.operatorId);
        return jsonToolResult({
          connections: connections
            .filter((item) => item.status !== "disconnected")
            .map(publicConnectionView),
        });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_microsoft_connection",
    {
      title: "Get Microsoft connection",
      description:
        "Returns one Microsoft connection owned by the authenticated operator. Requires connectionId.",
      inputSchema: z.object({ connectionId: connectionIdSchema }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const connection = await resolveOwnedConnection(operator.operatorId, input.connectionId);
        return jsonToolResult(publicConnectionView(connection));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "start_microsoft_connection",
    {
      title: "Connect another Microsoft account",
      description:
        "Creates a time-limited URL that starts an independent Microsoft OAuth connection for the authenticated operator. Pass connectionId to reauthorize an existing connection. Open the URL in a browser. Never returns OAuth secrets.",
      inputSchema: z.object({
        connectionId: connectionIdSchema.optional(),
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        if (input.connectionId) {
          await resolveOwnedConnection(operator.operatorId, input.connectionId);
        }
        const token = issueConnectToken({
          operatorId: operator.operatorId,
          operation: input.connectionId ? "reconnect" : "add_connection",
          connectionId: input.connectionId,
        });
        return jsonToolResult({
          url: connectUrl(token),
          expiresInSeconds: 900,
        });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "disconnect_microsoft_connection",
    {
      title: "Disconnect Microsoft connection",
      description:
        "Disconnects a Microsoft connection owned by the authenticated operator. The refresh token is removed. The connection is not deleted so it can be reauthorized later.",
      inputSchema: z.object({ connectionId: connectionIdSchema }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const connection = await getAppStore().disconnectConnection(
          operator.operatorId,
          input.connectionId,
        );
        return jsonToolResult(publicConnectionView(connection));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
