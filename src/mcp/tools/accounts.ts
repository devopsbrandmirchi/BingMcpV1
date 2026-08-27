import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  optionalMicrosoftIdSchema,
  optionalRefreshSchema,
} from "@/mcp/tools/schemas";
import { discoverAndPersistResources } from "@/services/discovery";
import {
  loadCurrentOperatorRecord,
  resolveAccountAccess,
  resolveOwnedConnection,
} from "@/services/resolver";
import { getAccessTokenForConnection } from "@/services/tokenService";
import { getAppStore } from "@/store/app-store";

function publicAccount(account: {
  accountId: string;
  customerId: string;
  accountName: string;
  accountNumber: string | null;
  status: string | null;
  currencyCode: string | null;
  timeZone: string | null;
  accountType: string | null;
  connectionId: string;
}) {
  return {
    accountId: account.accountId,
    customerId: account.customerId,
    accountName: account.accountName,
    accountNumber: account.accountNumber,
    status: account.status,
    currencyCode: account.currencyCode,
    timeZone: account.timeZone,
    accountType: account.accountType,
    connectionId: account.connectionId,
  };
}

export function registerAccountTools(server: McpServer): void {
  server.registerTool(
    "list_microsoft_accounts",
    {
      title: "List Microsoft Advertising accounts",
      description:
        "Lists advertising accounts accessible through the authenticated operator. Filter with connectionId, customerId, or accountId. Set refresh=true to rediscover from Microsoft.",
      inputSchema: z.object({
        connectionId: optionalConnectionIdSchema,
        customerId: optionalMicrosoftIdSchema,
        accountId: optionalMicrosoftIdSchema,
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
        let accounts = await getAppStore().listAccounts(operator.operatorId, input.connectionId);
        if (input.customerId) {
          accounts = accounts.filter((item) => item.customerId === input.customerId);
        }
        if (input.accountId) {
          accounts = accounts.filter((item) => item.accountId === input.accountId);
        }
        return jsonToolResult({ accounts: accounts.map(publicAccount) });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_microsoft_account",
    {
      title: "Get Microsoft Advertising account",
      description:
        "Returns one advertising account owned by the authenticated operator. Requires accountId. Specify connectionId if the same account is visible through more than one connection.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        customerId: optionalMicrosoftIdSchema,
        connectionId: optionalConnectionIdSchema,
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const resolved = await resolveAccountAccess({
          operatorId: operator.operatorId,
          accountId: input.accountId,
          customerId: input.customerId,
          connectionId: input.connectionId,
        });
        return jsonToolResult(publicAccount(resolved.account));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
