import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText } from "@/lib/errors";
import type { ReportTimePeriod } from "@/lib/dates";
import { runPerformanceReport, type PerformanceReportType } from "@/microsoft/reporting/api";
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
}
