import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText, ValidationError } from "@/lib/errors";
import {
  getCampaignsByAccountId,
  getCampaignsByIds,
} from "@/microsoft/campaign-management/api";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  pagingSchema,
} from "@/mcp/tools/schemas";
import { applyPaging, loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

export function registerCampaignTools(server: McpServer): void {
  server.registerTool(
    "list_campaigns",
    {
      title: "List campaigns",
      description:
        "Lists campaigns for a Microsoft Advertising account. Requires accountId. Specify connectionId if the account is accessible through more than one Microsoft connection.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        customerId: microsoftIdSchema.optional(),
        connectionId: optionalConnectionIdSchema,
        ...pagingSchema,
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const access = await resolveAccountAccess({
          operatorId: operator.operatorId,
          accountId: input.accountId,
          customerId: input.customerId,
          connectionId: input.connectionId,
        });
        const campaigns = await getCampaignsByAccountId({
          accessToken: access.accessToken,
          customerId: access.account.customerId,
          accountId: access.account.accountId,
        });
        return jsonToolResult(applyPaging(campaigns, input.limit, input.offset));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_campaign",
    {
      title: "Get campaign",
      description:
        "Returns one campaign by campaignId. Requires accountId and campaignId. Specify connectionId if the account is ambiguous.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        campaignId: microsoftIdSchema,
        customerId: microsoftIdSchema.optional(),
        connectionId: optionalConnectionIdSchema,
      }),
    },
    async (input) => {
      try {
        const operator = await loadCurrentOperatorRecord();
        const access = await resolveAccountAccess({
          operatorId: operator.operatorId,
          accountId: input.accountId,
          customerId: input.customerId,
          connectionId: input.connectionId,
        });
        const campaigns = await getCampaignsByIds(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          [input.campaignId],
        );
        const campaign = campaigns[0];
        if (!campaign) {
          throw new ValidationError("The requested campaign was not found.");
        }
        return jsonToolResult(campaign);
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
