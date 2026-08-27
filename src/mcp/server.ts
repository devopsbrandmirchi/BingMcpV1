import { createMcpHandler } from "mcp-handler";
import { APP_VERSION } from "@/lib/config";
import { logger } from "@/lib/logger";
import { createRequestId, runWithOperator } from "@/lib/request-context";
import { extractMcpToken, tryReadAccessToken } from "@/mcp/auth";
import {
  mcpOptionsResponse,
  mcpProbeGetResponse,
  toJsonRpcResponse,
  withCors,
  withStreamableAccept,
} from "@/mcp/http";
import { wwwAuthenticateHeader } from "@/mcp/oauth/metadata";
import { registerAccountTools } from "@/mcp/tools/accounts";
import { registerAdGroupTools } from "@/mcp/tools/adgroups";
import { registerAdTools } from "@/mcp/tools/ads";
import { registerCampaignTools } from "@/mcp/tools/campaigns";
import { registerConnectionTools } from "@/mcp/tools/connections";
import { registerCustomerTools } from "@/mcp/tools/customers";
import { registerKeywordTools } from "@/mcp/tools/keywords";
import { registerGetOperatorTool } from "@/mcp/tools/operator";
import { registerReportTools } from "@/mcp/tools/reports";
import { getAppStore } from "@/store/app-store";
import { normalizeSessionId } from "@/store/types";

export function unauthorizedMcpResponse(): Response {
  return withCors(
    new Response(
      JSON.stringify({
        error: "invalid_token",
        error_description: "Authentication required",
      }),
      {
        status: 401,
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
          "WWW-Authenticate": wwwAuthenticateHeader(),
        },
      },
    ),
  );
}

function createInnerHandler() {
  return createMcpHandler(
    (server) => {
      registerGetOperatorTool(server);
      registerConnectionTools(server);
      registerCustomerTools(server);
      registerAccountTools(server);
      registerCampaignTools(server);
      registerAdGroupTools(server);
      registerAdTools(server);
      registerKeywordTools(server);
      registerReportTools(server);
    },
    {
      serverInfo: {
        name: "bing-mcp-v1",
        version: APP_VERSION,
      },
      instructions:
        "This is a standalone Microsoft Advertising MCP connector. Identify the operator with get_operator, list Microsoft connections with list_microsoft_connections, then list customers and accounts. Use list_campaigns, list_ad_groups, list_ads, and list_keywords for entity reads. Use get_*_performance tools for official Reporting API data. Specify connectionId when the same advertising account is visible through more than one Microsoft connection. Never invent IDs. V1 is read-only.",
      onEvent: (event) => {
        if (event.type === "ERROR") {
          logger.error("MCP handler error", {
            source: event.source,
            severity: event.severity,
            context: event.context,
          });
        }
      },
    },
  );
}

function toolNameFromBody(body: unknown): string | undefined {
  const messages = Array.isArray(body) ? body : [body];
  for (const message of messages) {
    if (message && typeof message === "object") {
      const params = (message as { params?: { name?: unknown } }).params;
      if (typeof params?.name === "string") {
        return params.name;
      }
    }
  }
  return undefined;
}

export function createBingMcpHandler() {
  const handler = createInnerHandler();

  return async (req: Request): Promise<Response> => {
    const requestId = createRequestId(req);
    const started = Date.now();

    if (req.method === "OPTIONS") {
      return mcpOptionsResponse();
    }

    let body: unknown;
    if (req.method === "POST") {
      try {
        body = await req.clone().json();
      } catch {
        body = undefined;
      }
    }
    const mcpOperation = toolNameFromBody(body) ?? req.method.toLowerCase();

    const token = extractMcpToken(req);
    const payload = tryReadAccessToken(token);
    if (!payload?.sub) {
      logger.warn("Unauthorized MCP request", { requestId, mcpOperation, success: false });
      return unauthorizedMcpResponse();
    }

    const operator = await getAppStore().getOperator(payload.sub);
    if (!operator) {
      logger.warn("Unknown operator for MCP token", { requestId, mcpOperation, success: false });
      return unauthorizedMcpResponse();
    }

    void getAppStore().touchLastAccess(operator.operatorId);

    if (req.method === "GET") {
      return mcpProbeGetResponse();
    }
    if (req.method === "DELETE") {
      return new Response(null, { status: 204 });
    }

    return runWithOperator(
      {
        requestId,
        operatorId: operator.operatorId,
        sessionId: normalizeSessionId(payload.sid),
      },
      async () => {
        try {
          const response = await toJsonRpcResponse(await handler(withStreamableAccept(req)));
          logger.info("MCP request", {
            requestId,
            operatorId: operator.operatorId,
            mcpOperation,
            success: response.ok,
            durationMs: Date.now() - started,
          });
          return response;
        } catch (error) {
          logger.error("MCP request failed", {
            requestId,
            operatorId: operator.operatorId,
            mcpOperation,
            success: false,
            errorCategory: error instanceof Error ? error.name : "unknown",
            durationMs: Date.now() - started,
          });
          throw error;
        }
      },
    );
  };
}
