import type { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { toToolErrorText, ValidationError } from "@/lib/errors";
import {
  getConversionGoalsByAccount,
  getConversionGoalsByIds,
  getUetTagsByAccount,
} from "@/microsoft/campaign-management/api";
import {
  jsonToolResult,
  microsoftIdSchema,
  optionalConnectionIdSchema,
  pagingSchema,
} from "@/mcp/tools/schemas";
import { applyPaging, loadCurrentOperatorRecord, resolveAccountAccess } from "@/services/resolver";

export function registerConversionGoalTools(server: McpServer): void {
  server.registerTool(
    "list_conversion_goals",
    {
      title: "List conversion goals",
      description:
        "Lists conversion goals for a Microsoft Advertising account, plus the customer UET tags those goals can use. Requires accountId. Specify connectionId if the account is accessible through more than one Microsoft connection. Does not return UET tracking scripts.",
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
        const context = {
          accessToken: access.accessToken,
          customerId: access.account.customerId,
          accountId: access.account.accountId,
        };
        const [conversionGoals, uetTags] = await Promise.all([
          getConversionGoalsByAccount(context),
          getUetTagsByAccount(context),
        ]);
        return jsonToolResult({
          ...applyPaging(conversionGoals, input.limit, input.offset),
          uetTags,
        });
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "get_conversion_goal",
    {
      title: "Get conversion goal",
      description:
        "Returns one conversion goal by conversionGoalId. Requires accountId and conversionGoalId. Specify connectionId if the account is ambiguous.",
      inputSchema: z.object({
        accountId: microsoftIdSchema,
        conversionGoalId: microsoftIdSchema,
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
        const goals = await getConversionGoalsByIds(
          {
            accessToken: access.accessToken,
            customerId: access.account.customerId,
            accountId: access.account.accountId,
          },
          [input.conversionGoalId],
        );
        const goal = goals[0];
        if (!goal) {
          throw new ValidationError("The requested conversion goal was not found.");
        }
        return jsonToolResult(goal);
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );

  server.registerTool(
    "list_uet_tags",
    {
      title: "List UET tags",
      description:
        "Lists Universal Event Tracking tags for the Microsoft Advertising customer that owns the account. Requires accountId. Does not return tracking scripts.",
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
        const tags = await getUetTagsByAccount({
          accessToken: access.accessToken,
          customerId: access.account.customerId,
          accountId: access.account.accountId,
        });
        return jsonToolResult(applyPaging(tags, input.limit, input.offset));
      } catch (error) {
        return { content: [{ type: "text", text: toToolErrorText(error) }], isError: true };
      }
    },
  );
}
