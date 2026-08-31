import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import type { ReportTimePeriod } from "@/lib/dates";
import {
  rankSearchQueryRows,
  runPerformanceReport,
  SEARCH_QUERY_SORT_FIELDS,
  type PerformanceReportType,
} from "@/microsoft/reporting/api";
import {
  isoDateSchema,
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  reportPeriodSchema,
} from "@/mcp/tools/schemas";
import { loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

const reportBase = {
  accountId: microsoftIdSchema,
  customerId: microsoftIdSchema.optional(),
  connectionId: optionalConnectionIdSchema,
  period: reportPeriodSchema.optional(),
  startDate: isoDateSchema.optional(),
  endDate: isoDateSchema.optional(),
};

async function runReport(
  type: PerformanceReportType,
  input: {
    accountId: string;
    customerId?: string;
    connectionId?: string;
    period?: ReportTimePeriod;
    startDate?: string;
    endDate?: string;
    campaignIds?: string[];
    adGroupIds?: string[];
  },
) {
  const operator = await loadCurrentOperatorRecord();
  const access = await resolveAccountAccess({
    operatorId: operator.operatorId,
    accountId: input.accountId,
    customerId: input.customerId,
    connectionId: input.connectionId,
  });
  return runPerformanceReport({
    context: {
      accessToken: access.accessToken,
      customerId: access.account.customerId,
      accountId: access.account.accountId,
    },
    type,
    period: input.period,
    startDate: input.startDate,
    endDate: input.endDate,
    campaignIds: input.campaignIds,
    adGroupIds: input.adGroupIds,
  });
}

export function registerReportTools(server: McpServer): void {
  server.registerTool(
    "get_account_performance",
    {
      title: "Account performance report",
      description:
        "Returns daily Microsoft Advertising account performance. Prefer period=ThisMonth or ThisWeek. Otherwise provide startDate and endDate as YYYY-MM-DD. Requires accountId.",
      inputSchema: z.object(reportBase),
    },
    async (input) => {
      try {
        return jsonToolResult(await runReport("AccountPerformanceReportRequest", input));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_campaign_performance",
    {
      title: "Campaign performance report",
      description:
        "Returns daily campaign performance. Prefer period=ThisMonth or ThisWeek. Otherwise provide startDate and endDate as YYYY-MM-DD. Requires accountId. Optional campaignId filter.",
      inputSchema: z.object({
        ...reportBase,
        campaignId: microsoftIdSchema.optional(),
      }),
    },
    async (input) => {
      try {
        return jsonToolResult(
          await runReport("CampaignPerformanceReportRequest", {
            ...input,
            campaignIds: input.campaignId ? [input.campaignId] : undefined,
          }),
        );
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_ad_group_performance",
    {
      title: "Ad group performance report",
      description:
        "Returns daily ad group performance. Prefer period=ThisMonth or ThisWeek. Otherwise provide startDate and endDate as YYYY-MM-DD. Requires accountId. Optional campaignId or adGroupId filters.",
      inputSchema: z.object({
        ...reportBase,
        campaignId: microsoftIdSchema.optional(),
        adGroupId: microsoftIdSchema.optional(),
      }),
    },
    async (input) => {
      try {
        return jsonToolResult(
          await runReport("AdGroupPerformanceReportRequest", {
            ...input,
            campaignIds: input.campaignId ? [input.campaignId] : undefined,
            adGroupIds: input.adGroupId ? [input.adGroupId] : undefined,
          }),
        );
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_keyword_performance",
    {
      title: "Keyword performance report",
      description:
        "Returns daily keyword performance. Prefer period=ThisMonth or ThisWeek. Otherwise provide startDate and endDate as YYYY-MM-DD. Requires accountId. Optional campaignId or adGroupId filters.",
      inputSchema: z.object({
        ...reportBase,
        campaignId: microsoftIdSchema.optional(),
        adGroupId: microsoftIdSchema.optional(),
      }),
    },
    async (input) => {
      try {
        return jsonToolResult(
          await runReport("KeywordPerformanceReportRequest", {
            ...input,
            campaignIds: input.campaignId ? [input.campaignId] : undefined,
            adGroupIds: input.adGroupId ? [input.adGroupId] : undefined,
          }),
        );
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_search_query_performance",
    {
      title: "Search query performance report",
      description:
        "Returns actual user search queries (search terms) that triggered ads, not the keywords you bid on. Uses Microsoft SearchQueryPerformanceReport. Prefer period=ThisMonth or ThisWeek, or startDate and endDate as YYYY-MM-DD. Requires accountId. Optional campaignId or adGroupId. Default sort is spend descending. Microsoft only returns queries with a significant number of clicks in the last 30 days, and search-query data applies to Search/search-network campaigns (not Performance Max). Current-day data may be incomplete.",
      inputSchema: z.object({
        ...reportBase,
        campaignId: microsoftIdSchema.optional(),
        adGroupId: microsoftIdSchema.optional(),
        sortBy: z.enum(SEARCH_QUERY_SORT_FIELDS).default("spend"),
        limit: z.number().int().min(1).max(1000).default(20),
        minClicks: z.number().int().min(0).optional(),
        maxConversions: z.number().min(0).optional(),
      }),
    },
    async (input) => {
      try {
        const report = await runReport("SearchQueryPerformanceReportRequest", {
          ...input,
          campaignIds: input.campaignId ? [input.campaignId] : undefined,
          adGroupIds: input.adGroupId ? [input.adGroupId] : undefined,
        });
        const rows = rankSearchQueryRows(report.rows, {
          sortBy: input.sortBy,
          limit: input.limit,
          minClicks: input.minClicks,
          maxConversions: input.maxConversions,
        });
        return jsonToolResult({
          ...report,
          rows,
          rowCount: rows.length,
          truncated: report.truncated || report.rows.length > rows.length,
          sortBy: input.sortBy,
        });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
