import { microsoftSoapRequest } from "@/microsoft/client/soap";
import { childBlocks, childText, longArrayXml, stringArrayXml, stripXmlNamespaces } from "@/microsoft/client/xml";
import {
  CAMPAIGN_MANAGEMENT_NAMESPACE,
  CAMPAIGN_MANAGEMENT_SOAP_URL,
} from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { Ad, AdGroup, Campaign, Keyword } from "@/microsoft/models/types";

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

function asString(value: string | null | undefined): string | null {
  return value && value.length > 0 ? value : null;
}

function asNumber(value: string | null | undefined): number | null {
  if (!value) {
    return null;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

async function campaignSoap(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  action: string,
  innerXml: string,
): Promise<string> {
  const xml = await microsoftSoapRequest({
    service: "campaign-management",
    url: CAMPAIGN_MANAGEMENT_SOAP_URL,
    action,
    namespace: CAMPAIGN_MANAGEMENT_NAMESPACE,
    context,
    bodyXml: `<${action}Request xmlns="${CAMPAIGN_MANAGEMENT_NAMESPACE}">${innerXml}</${action}Request>`,
  });
  return stripXmlNamespaces(xml);
}

function mapCampaign(
  block: string,
  context: { accountId: string; customerId: string },
): Campaign | null {
  const campaignId = childText(block, "Id");
  if (!campaignId) {
    return null;
  }
  return {
    campaignId,
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(childText(block, "Name")),
    status: asString(childText(block, "Status")),
    campaignType: asString(childText(block, "CampaignType")),
    budget: asNumber(childText(block, "DailyBudget") ?? childText(block, "Budget")),
    budgetType: asString(childText(block, "BudgetType")),
    timeZone: asString(childText(block, "TimeZone")),
  };
}

function mapAdGroup(
  block: string,
  context: { accountId: string; customerId: string; campaignId?: string },
): AdGroup | null {
  const adGroupId = childText(block, "Id");
  if (!adGroupId) {
    return null;
  }
  return {
    adGroupId,
    campaignId: asString(childText(block, "CampaignId")) ?? context.campaignId ?? "",
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(childText(block, "Name")),
    status: asString(childText(block, "Status")),
  };
}

function mapAd(
  block: string,
  context: { accountId: string; customerId: string; adGroupId?: string },
): Ad | null {
  const adId = childText(block, "Id");
  if (!adId) {
    return null;
  }
  const finalUrlsXml = childText(block, "FinalUrls");
  return {
    adId,
    adGroupId: asString(childText(block, "AdGroupId")) ?? context.adGroupId ?? "",
    accountId: context.accountId,
    customerId: context.customerId,
    type: asString(childText(block, "Type")),
    status: asString(childText(block, "Status")),
    editorialStatus: asString(childText(block, "EditorialStatus")),
    headline: asString(childText(block, "Title") ?? childText(block, "Headline")),
    text: asString(childText(block, "Text")),
    displayUrl: asString(childText(block, "DisplayUrl")),
    finalUrls: finalUrlsXml ? childBlocks(finalUrlsXml, "string").map((item) => item.trim()).filter(Boolean) : [],
  };
}

function mapKeyword(
  block: string,
  context: { accountId: string; customerId: string; adGroupId?: string },
): Keyword | null {
  const keywordId = childText(block, "Id");
  if (!keywordId) {
    return null;
  }
  const bidXml = childText(block, "Bid");
  return {
    keywordId,
    adGroupId: asString(childText(block, "AdGroupId")) ?? context.adGroupId ?? "",
    accountId: context.accountId,
    customerId: context.customerId,
    text: asString(childText(block, "Text")),
    matchType: asString(childText(block, "MatchType")),
    status: asString(childText(block, "Status")),
    bid: asNumber(bidXml ? childText(bidXml, "Amount") : null) ?? asNumber(bidXml),
  };
}

export async function getCampaignsByAccountId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
): Promise<Campaign[]> {
  const xml = await campaignSoap(
    context,
    "GetCampaignsByAccountId",
    `<AccountId>${context.accountId}</AccountId><CampaignType>${ALL_CAMPAIGN_TYPES}</CampaignType>`,
  );
  return childBlocks(xml, "Campaign")
    .map((block) => mapCampaign(block, context))
    .filter((item): item is Campaign => Boolean(item));
}

export async function getCampaignsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  campaignIds: string[],
): Promise<Campaign[]> {
  const xml = await campaignSoap(
    context,
    "GetCampaignsByIds",
    `<AccountId>${context.accountId}</AccountId>${longArrayXml("CampaignIds", campaignIds)}<CampaignType>${ALL_CAMPAIGN_TYPES}</CampaignType>`,
  );
  return childBlocks(xml, "Campaign")
    .map((block) => mapCampaign(block, context))
    .filter((item): item is Campaign => Boolean(item));
}

export async function getAdGroupsByCampaignId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  campaignId: string,
): Promise<AdGroup[]> {
  const xml = await campaignSoap(
    context,
    "GetAdGroupsByCampaignId",
    `<CampaignId>${campaignId}</CampaignId>`,
  );
  return childBlocks(xml, "AdGroup")
    .map((block) => mapAdGroup(block, { ...context, campaignId }))
    .filter((item): item is AdGroup => Boolean(item));
}

export async function getAdGroupsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupIds: string[],
): Promise<AdGroup[]> {
  const xml = await campaignSoap(
    context,
    "GetAdGroupsByIds",
    longArrayXml("AdGroupIds", adGroupIds),
  );
  return childBlocks(xml, "AdGroup")
    .map((block) => mapAdGroup(block, context))
    .filter((item): item is AdGroup => Boolean(item));
}

export async function getAdsByAdGroupId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupId: string,
): Promise<Ad[]> {
  const xml = await campaignSoap(
    context,
    "GetAdsByAdGroupId",
    `<AdGroupId>${adGroupId}</AdGroupId>${stringArrayXml("AdTypes", "AdType", ALL_AD_TYPES)}`,
  );
  return childBlocks(xml, "Ad")
    .map((block) => mapAd(block, { ...context, adGroupId }))
    .filter((item): item is Ad => Boolean(item));
}

export async function getAdsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adIds: string[],
): Promise<Ad[]> {
  const xml = await campaignSoap(
    context,
    "GetAdsByIds",
    `${longArrayXml("AdIds", adIds)}${stringArrayXml("AdTypes", "AdType", ALL_AD_TYPES)}`,
  );
  return childBlocks(xml, "Ad")
    .map((block) => mapAd(block, context))
    .filter((item): item is Ad => Boolean(item));
}

export async function getKeywordsByAdGroupId(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  adGroupId: string,
): Promise<Keyword[]> {
  const xml = await campaignSoap(
    context,
    "GetKeywordsByAdGroupId",
    `<AdGroupId>${adGroupId}</AdGroupId>`,
  );
  return childBlocks(xml, "Keyword")
    .map((block) => mapKeyword(block, { ...context, adGroupId }))
    .filter((item): item is Keyword => Boolean(item));
}

export async function getKeywordsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  keywordIds: string[],
): Promise<Keyword[]> {
  const xml = await campaignSoap(
    context,
    "GetKeywordsByIds",
    longArrayXml("KeywordIds", keywordIds),
  );
  return childBlocks(xml, "Keyword")
    .map((block) => mapKeyword(block, context))
    .filter((item): item is Keyword => Boolean(item));
}
