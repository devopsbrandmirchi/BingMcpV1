import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText, ValidationError } from "@/lib/errors";
import {
  getAdGroupsByCampaignId,
  getAdsByAdGroupId,
  getAdsByIds,
} from "@/microsoft/campaign-management/api";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  pagingSchema,
} from "@/mcp/tools/schemas";
import { applyPaging, loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

export function registerAdTools(server: McpServer): void {
  server.registerTool(
    "list_ads",
    {
      title: "List ads",
      description:
        "Lists ads. Requires accountId and either adGroupId or campaignId. When campaignId is supplied, ads are collected from ad groups in that campaign up to the paging limit.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        adGroupId: microsoftIdSchema.optional(),
        campaignId: microsoftIdSchema.optional(),
        customerId: microsoftIdSchema.optional(),
        connectionId: optionalConnectionIdSchema,
        status: z.string().min(1).optional(),
        ...pagingSchema,
      }),
    },
    async (input) => {
      try {
        if (!input.adGroupId && !input.campaignId) {
          throw new ValidationError("Provide adGroupId or campaignId.");
        }
        const operator = await loadCurrentOperatorRecord();
        const access = await resolveAccountAccess({
          operatorId: operator.operatorId,
          accountId: input.accountId,
          customerId: input.customerId,
          connectionId: input.connectionId,
        });
        const context = {
          accessToken: access.accessToken,
          customerId: access.account.customerId,
          accountId: access.account.accountId,
        };
        let ads = input.adGroupId
          ? await getAdsByAdGroupId(context, input.adGroupId)
          : [];
        if (!input.adGroupId && input.campaignId) {
          const adGroups = await getAdGroupsByCampaignId(context, input.campaignId);
          for (const adGroup of adGroups) {
            ads = ads.concat(await getAdsByAdGroupId(context, adGroup.adGroupId));
            if (ads.length >= input.offset + input.limit) {
              break;
            }
          }
        }
        if (input.status) {
          ads = ads.filter((ad) => ad.status === input.status);
        }
        return jsonToolResult(applyPaging(ads, input.limit, input.offset));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_ad",
    {
      title: "Get ad",
      description: "Returns one ad by adId. Requires accountId and adId.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        adId: microsoftIdSchema,
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
        const ads = await getAdsByIds(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          [input.adId],
        );
        const ad = ads[0];
        if (!ad) {
          throw new ValidationError("The requested ad was not found.");
        }
        return jsonToolResult(ad);
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
