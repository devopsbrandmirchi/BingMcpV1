import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText, ValidationError } from "@/lib/errors";
import {
  getAdGroupsByCampaignId,
  getKeywordsByAdGroupId,
  getKeywordsByIds,
} from "@/microsoft/campaign-management/api";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  pagingSchema,
} from "@/mcp/tools/schemas";
import { applyPaging, loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

export function registerKeywordTools(server: McpServer): void {
  server.registerTool(
    "list_keywords",
    {
      title: "List keywords",
      description:
        "Lists keywords. Requires accountId and either adGroupId or campaignId. When campaignId is supplied, keywords are collected from ad groups in that campaign up to the paging limit.",
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
        let keywords = input.adGroupId
          ? await getKeywordsByAdGroupId(context, input.adGroupId)
          : [];
        if (!input.adGroupId && input.campaignId) {
          const adGroups = await getAdGroupsByCampaignId(context, input.campaignId);
          for (const adGroup of adGroups) {
            keywords = keywords.concat(await getKeywordsByAdGroupId(context, adGroup.adGroupId));
            if (keywords.length >= input.offset + input.limit) {
              break;
            }
          }
        }
        if (input.status) {
          keywords = keywords.filter((keyword) => keyword.status === input.status);
        }
        return jsonToolResult(applyPaging(keywords, input.limit, input.offset));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_keyword",
    {
      title: "Get keyword",
      description: "Returns one keyword by keywordId. Requires accountId and keywordId.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        keywordId: microsoftIdSchema,
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
        const keywords = await getKeywordsByIds(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          [input.keywordId],
        );
        const keyword = keywords[0];
        if (!keyword) {
          throw new ValidationError("The requested keyword was not found.");
        }
        return jsonToolResult(keyword);
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
