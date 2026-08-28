import { microsoftJsonRequest } from "@/microsoft/client/http";
import { toMicrosoftLong, toMicrosoftLongs } from "@/microsoft/client/ids";
import { CAMPAIGN_MANAGEMENT_BASE } from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { Ad, AdGroup, Campaign, Keyword } from "@/microsoft/models/types";

// REST CampaignType is a flags string, not a JSON array.
const ALL_CAMPAIGN_TYPES = "Search Shopping DynamicSearchAds Audience Hotel PerformanceMax App";

const ALL_AD_TYPES = [
  "Text",
  "ExpandedText",
  "ResponsiveSearch",
  "ResponsiveAd",
  "Image",
  "Product",
  "Hotel",
  "AppInstall",
  "DynamicSearch",
];

function asString(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null;
  }
  return String(value);
}

function asNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function asStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((item) => String(item)).filter(Boolean);
}

export async function getCampaignsByAccountId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
): Promise<Campaign[]> {
  const response = await microsoftJsonRequest<{ Campaigns?: Array<Record<string, unknown>> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Campaigns/QueryByAccountId`,
    context,
    body: {
      AccountId: toMicrosoftLong(context.accountId, "accountId"),
      CampaignType: ALL_CAMPAIGN_TYPES,
    },
  });

  return (response.Campaigns ?? []).map((item) => ({
    campaignId: String(item.Id ?? ""),
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(item.Name),
    status: asString(item.Status),
    campaignType: asString(item.CampaignType),
    budget: asNumber(item.DailyBudget ?? item.Budget),
    budgetType: asString(item.BudgetType),
    timeZone: asString(item.TimeZone),
  })).filter((item) => item.campaignId);
}

export async function getCampaignsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  campaignIds: string[],
): Promise<Campaign[]> {
  const response = await microsoftJsonRequest<{ Campaigns?: Array<Record<string, unknown> | null> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Campaigns/QueryByIds`,
    context,
    body: {
      AccountId: toMicrosoftLong(context.accountId, "accountId"),
      CampaignIds: toMicrosoftLongs(campaignIds, "campaignId"),
      CampaignType: ALL_CAMPAIGN_TYPES,
    },
  });
  return (response.Campaigns ?? [])
    .filter((item): item is Record<string, unknown> => Boolean(item))
    .map((item) => ({
      campaignId: String(item.Id ?? ""),
      accountId: context.accountId,
      customerId: context.customerId,
      name: asString(item.Name),
      status: asString(item.Status),
      campaignType: asString(item.CampaignType),
      budget: asNumber(item.DailyBudget ?? item.Budget),
      budgetType: asString(item.BudgetType),
      timeZone: asString(item.TimeZone),
    }))
    .filter((item) => item.campaignId);
}

export async function getAdGroupsByCampaignId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  campaignId: string,
): Promise<AdGroup[]> {
  const response = await microsoftJsonRequest<{ AdGroups?: Array<Record<string, unknown>> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/AdGroups/QueryByCampaignId`,
    context,
    body: { CampaignId: toMicrosoftLong(campaignId, "campaignId") },
  });
  return (response.AdGroups ?? []).map((item) => ({
    adGroupId: String(item.Id ?? ""),
    campaignId,
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(item.Name),
    status: asString(item.Status),
  })).filter((item) => item.adGroupId);
}

export async function getAdGroupsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupIds: string[],
): Promise<AdGroup[]> {
  const response = await microsoftJsonRequest<{ AdGroups?: Array<Record<string, unknown> | null> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/AdGroups/QueryByIds`,
    context,
    body: { AdGroupIds: toMicrosoftLongs(adGroupIds, "adGroupId") },
  });
  return (response.AdGroups ?? [])
    .filter((item): item is Record<string, unknown> => Boolean(item))
    .map((item) => ({
      adGroupId: String(item.Id ?? ""),
      campaignId: asString(item.CampaignId) ?? "",
      accountId: context.accountId,
      customerId: context.customerId,
      name: asString(item.Name),
      status: asString(item.Status),
    }))
    .filter((item) => item.adGroupId);
}

export async function getAdsByAdGroupId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupId: string,
): Promise<Ad[]> {
  const response = await microsoftJsonRequest<{ Ads?: Array<Record<string, unknown>> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Ads/QueryByAdGroupId`,
    context,
    body: {
      AdGroupId: toMicrosoftLong(adGroupId, "adGroupId"),
      AdTypes: ALL_AD_TYPES,
    },
  });
  return (response.Ads ?? []).map((item) => ({
    adId: String(item.Id ?? ""),
    adGroupId,
    accountId: context.accountId,
    customerId: context.customerId,
    type: asString(item.Type),
    status: asString(item.Status),
    editorialStatus: asString(item.EditorialStatus),
    headline: asString(item.Title ?? item.Headline),
    text: asString(item.Text),
    displayUrl: asString(item.DisplayUrl),
    finalUrls: asStringArray(item.FinalUrls),
  })).filter((item) => item.adId);
}

export async function getAdsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adIds: string[],
): Promise<Ad[]> {
  const response = await microsoftJsonRequest<{ Ads?: Array<Record<string, unknown> | null> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Ads/QueryByIds`,
    context,
    body: {
      AdIds: toMicrosoftLongs(adIds, "adId"),
      AdTypes: ALL_AD_TYPES,
    },
  });
  return (response.Ads ?? [])
    .filter((item): item is Record<string, unknown> => Boolean(item))
    .map((item) => ({
      adId: String(item.Id ?? ""),
      adGroupId: asString(item.AdGroupId) ?? "",
      accountId: context.accountId,
      customerId: context.customerId,
      type: asString(item.Type),
      status: asString(item.Status),
      editorialStatus: asString(item.EditorialStatus),
      headline: asString(item.Title ?? item.Headline),
      text: asString(item.Text),
      displayUrl: asString(item.DisplayUrl),
      finalUrls: asStringArray(item.FinalUrls),
    }))
    .filter((item) => item.adId);
}

export async function getKeywordsByAdGroupId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupId: string,
): Promise<Keyword[]> {
  const response = await microsoftJsonRequest<{ Keywords?: Array<Record<string, unknown>> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Keywords/QueryByAdGroupId`,
    context,
    body: { AdGroupId: toMicrosoftLong(adGroupId, "adGroupId") },
  });
  return (response.Keywords ?? []).map((item) => ({
    keywordId: String(item.Id ?? ""),
    adGroupId,
    accountId: context.accountId,
    customerId: context.customerId,
    text: asString(item.Text),
    matchType: asString(item.MatchType),
    status: asString(item.Status),
    bid: asNumber((item.Bid as { Amount?: unknown } | undefined)?.Amount ?? item.Bid),
  })).filter((item) => item.keywordId);
}

export async function getKeywordsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  keywordIds: string[],
): Promise<Keyword[]> {
  const response = await microsoftJsonRequest<{ Keywords?: Array<Record<string, unknown> | null> }>({
    service: "campaign-management",
    url: `${CAMPAIGN_MANAGEMENT_BASE}/Keywords/QueryByIds`,
    context,
    body: { KeywordIds: toMicrosoftLongs(keywordIds, "keywordId") },
  });
  return (response.Keywords ?? [])
    .filter((item): item is Record<string, unknown> => Boolean(item))
    .map((item) => ({
      keywordId: String(item.Id ?? ""),
      adGroupId: asString(item.AdGroupId) ?? "",
      accountId: context.accountId,
      customerId: context.customerId,
      text: asString(item.Text),
      matchType: asString(item.MatchType),
      status: asString(item.Status),
      bid: asNumber((item.Bid as { Amount?: unknown } | undefined)?.Amount ?? item.Bid),
    }))
    .filter((item) => item.keywordId);
}
