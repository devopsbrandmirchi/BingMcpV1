import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText, ValidationError } from "@/lib/errors";
import { getAdGroupsByCampaignId, getAdGroupsByIds } from "@/microsoft/campaign-management/api";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  pagingSchema,
} from "@/mcp/tools/schemas";
import { applyPaging, loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

export function registerAdGroupTools(server: McpServer): void {
  server.registerTool(
    "list_ad_groups",
    {
      title: "List ad groups",
      description:
        "Lists ad groups for a campaign. Requires accountId and campaignId. Specify connectionId if the account is ambiguous.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        campaignId: microsoftIdSchema,
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
        const adGroups = await getAdGroupsByCampaignId(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          input.campaignId,
        );
        return jsonToolResult(applyPaging(adGroups, input.limit, input.offset));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_ad_group",
    {
      title: "Get ad group",
      description: "Returns one ad group by adGroupId. Requires accountId and adGroupId.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        adGroupId: microsoftIdSchema,
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
        const adGroups = await getAdGroupsByIds(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          [input.adGroupId],
        );
        const adGroup = adGroups[0];
        if (!adGroup) {
          throw new ValidationError("The requested ad group was not found.");
        }
        return jsonToolResult(adGroup);
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
