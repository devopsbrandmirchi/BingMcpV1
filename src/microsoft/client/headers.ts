import { getConfig } from "@/lib/config";

export type MicrosoftService = "customer-management" | "campaign-management" | "reporting";

export interface MicrosoftRequestContext {
  accessToken: string;
  customerId?: string;
  accountId?: string;
}

export function buildMicrosoftHeaders(
  service: MicrosoftService,
  context: MicrosoftRequestContext,
): Record<string, string> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${context.accessToken}`,
    DeveloperToken: getConfig().microsoftAdsDeveloperToken,
    "Content-Type": "application/json",
    Accept: "application/json",
  };

  if (service === "campaign-management" || service === "reporting") {
    if (context.customerId) {
      headers.CustomerId = context.customerId;
    }
    if (context.accountId) {
      headers.CustomerAccountId = context.accountId;
    }
  }

  return headers;
}
