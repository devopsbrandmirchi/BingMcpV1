import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  optionalRefreshSchema,
} from "@/mcp/tools/schemas";
import { discoverAndPersistResources } from "@/services/discovery";
import {
  loadCurrentOperatorRecord,
  resolveCustomerAccess,
  resolveOwnedConnection,
} from "@/services/resolver";
import { getAccessTokenForConnection } from "@/services/tokenService";
import { getAppStore } from "@/store/app-store";

function publicCustomer(customer: {
  customerId: string;
  customerName: string;
  customerNumber: string | null;
  status: string | null;
  connectionId: string;
}) {
  return {
    customerId: customer.customerId,
    customerName: customer.customerName,
    customerNumber: customer.customerNumber,
    status: customer.status,
    connectionId: customer.connectionId,
  };
}

export function registerCustomerTools(server: McpServer): void {
  server.registerTool(
    "list_microsoft_customers",
    {
      title: "List Microsoft Advertising customers",
      description:
        "Lists Microsoft Advertising customers accessible through the authenticated operator's Microsoft connections. Optionally filter by connectionId. Set refresh=true to rediscover from Microsoft.",
      inputSchema: z.object({
        connectionId: optionalConnectionIdSchema,
        refresh: optionalRefreshSchema,
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        if (input.refresh) {
          const connections = input.connectionId
            ? [await resolveOwnedConnection(operator.operatorId, input.connectionId)]
            : (await getAppStore().listConnections(operator.operatorId)).filter(
                (item) => item.status === "active",
              );
          for (const connection of connections) {
            const accessToken = await getAccessTokenForConnection(
              operator.operatorId,
              connection.connectionId,
            );
            await discoverAndPersistResources(operator.operatorId, connection.connectionId, accessToken);
          }
        }
        const customers = await getAppStore().listCustomers(operator.operatorId, input.connectionId);
        return jsonToolResult({ customers: customers.map(publicCustomer) });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_microsoft_customer",
    {
      title: "Get Microsoft Advertising customer",
      description:
        "Returns one Microsoft Advertising customer owned by the authenticated operator. Requires customerId. Specify connectionId if the same customer is visible through more than one connection.",
      inputSchema: z.object({
        customerId: microsoftIdSchema,
        connectionId: optionalConnectionIdSchema,
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const resolved = await resolveCustomerAccess({
          operatorId: operator.operatorId,
          customerId: input.customerId,
          connectionId: input.connectionId,
        });
        return jsonToolResult(publicCustomer(resolved.customer));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
