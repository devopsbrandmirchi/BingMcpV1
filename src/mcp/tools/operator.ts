import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import { peekOperatorContext } from "@/lib/request-context";
import { jsonToolResult } from "@/mcp/tools/schemas";
import { loadCurrentOperatorRecord } from "@/services/resolver";
import { getAppStore } from "@/store/app-store";

export function registerGetOperatorTool(server: McpServer): void {
  server.registerTool(
    "get_operator",
    {
      title: "Get authenticated operator",
      description:
        "Returns the authenticated MCP operator for this Claude connector session. Use this to confirm operatorId before listing Microsoft connections. Never returns tokens or credentials.",
      inputSchema: z.object({}),
    },
    async () => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const context = peekOperatorContext();
        const connections = await getAppStore().listConnections(operator.operatorId);
        return jsonToolResult({
          operatorId: operator.operatorId,
          connectionCount: connections.filter((item) => item.status !== "disconnected").length,
          connected: connections.some((item) => item.status === "active"),
          requestId: context?.requestId,
        });
      } catch (error) {
        return {
          content: [{ type: "text", text: toToolErrorText(error) }],
          isError: true,
        };
      }
    },
  );
}
