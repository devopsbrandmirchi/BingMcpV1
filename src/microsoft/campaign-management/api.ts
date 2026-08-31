import { microsoftSoapRequest } from "@/microsoft/client/soap";
import { childBlocks, childText, longArrayXml, stringArrayXml, stripXmlNamespaces } from "@/microsoft/client/xml";
import {
  CAMPAIGN_MANAGEMENT_NAMESPACE,
  CAMPAIGN_MANAGEMENT_SOAP_URL,
} from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { Ad, AdGroup, Campaign, ConversionGoal, Keyword, UetTag } from "@/microsoft/models/types";

const ALL_CONVERSION_GOAL_TYPES =
  "Url Duration PagesViewedPerVisit Event AppInstall OfflineConversion InStoreTransaction AppDownload";

const CONVERSION_GOAL_ADDITIONAL_FIELDS =
  "AttributionModelType ViewThroughConversionWindowInMinutes";

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

function asBoolean(value: string | null | undefined): boolean | null {
  if (value === "true") {
    return true;
  }
  if (value === "false") {
    return false;
  }
  return null;
}

function conversionGoalBlocks(xml: string): Array<{ inner: string; typeHint: string | null }> {
  const regex = /<ConversionGoal(\s[^>]*)?>([\s\S]*?)<\/ConversionGoal>/gi;
  const blocks: Array<{ inner: string; typeHint: string | null }> = [];
  let match = regex.exec(xml);
  while (match) {
    const attrs = match[1] ?? "";
    const typeHint = attrs.match(/\b(?:i:)?type="([^"]+)"/i)?.[1] ?? null;
    blocks.push({ inner: match[2] ?? "", typeHint });
    match = regex.exec(xml);
  }
  return blocks;
}

function goalTypeFromHint(hint: string | null): string | null {
  if (!hint) {
    return null;
  }
  return hint.replace(/Goal$/i, "") || hint;
}

function mapConversionGoal(
  block: string,
  context: { accountId: string; customerId: string },
  typeHint?: string | null,
): ConversionGoal | null {
  const conversionGoalId = childText(block, "Id");
  if (!conversionGoalId) {
    return null;
  }
  const revenueBlock = childBlocks(block, "Revenue")[0];
  const withoutRevenue = block.replace(/<Revenue[\s\S]*?<\/Revenue>/gi, "");
  return {
    conversionGoalId,
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(childText(block, "Name")),
    type: goalTypeFromHint(typeHint ?? null) ?? asString(childText(withoutRevenue, "Type")),
    status: asString(childText(block, "Status")),
    trackingStatus: asString(childText(block, "TrackingStatus")),
    tagId: asString(childText(block, "TagId")),
    scope: asString(childText(block, "Scope")),
    countType: asString(childText(block, "CountType")),
    goalCategory: asString(childText(block, "GoalCategory")),
    excludeFromBidding: asBoolean(childText(block, "ExcludeFromBidding")),
    isAutoGoal: asBoolean(childText(block, "IsAutoGoal")),
    isEnhancedConversionsEnabled: asBoolean(childText(block, "IsEnhancedConversionsEnabled")),
    conversionWindowInMinutes: asNumber(childText(block, "ConversionWindowInMinutes")),
    viewThroughConversionWindowInMinutes: asNumber(childText(block, "ViewThroughConversionWindowInMinutes")),
    attributionModelType: asString(childText(block, "AttributionModelType")),
    revenue: revenueBlock
      ? {
          type: asString(childText(revenueBlock, "Type")),
          value: asNumber(childText(revenueBlock, "Value")),
          currencyCode: asString(childText(revenueBlock, "CurrencyCode")),
        }
      : null,
    urlExpression: asString(childText(block, "UrlExpression")),
    urlOperator: asString(childText(block, "UrlOperator")),
    actionExpression: asString(childText(block, "ActionExpression")),
    categoryExpression: asString(childText(block, "CategoryExpression")),
    labelExpression: asString(childText(block, "LabelExpression")),
    minimumDurationInSeconds: asNumber(childText(block, "MinimumDurationInSeconds")),
    minimumPagesViewed: asNumber(childText(block, "MinimumPagesViewed")),
  };
}

function mapUetTag(
  block: string,
  context: { accountId: string; customerId: string },
): UetTag | null {
  const uetTagId = childText(block, "Id");
  if (!uetTagId) {
    return null;
  }
  return {
    uetTagId,
    accountId: context.accountId,
    customerId: context.customerId,
    name: asString(childText(block, "Name")),
    description: asString(childText(block, "Description")),
    trackingStatus: asString(childText(block, "TrackingStatus")),
    industry: asString(childText(block, "Industry")),
  };
}

export async function getConversionGoalsByAccount(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
): Promise<ConversionGoal[]> {
  const xml = await campaignSoap(
    context,
    "GetConversionGoalsByIds",
    `<ConversionGoalIds i:nil="true"/><ConversionGoalTypes>${ALL_CONVERSION_GOAL_TYPES}</ConversionGoalTypes><ReturnAdditionalFields>${CONVERSION_GOAL_ADDITIONAL_FIELDS}</ReturnAdditionalFields>`,
  );
  return conversionGoalBlocks(xml)
    .map(({ inner, typeHint }) => mapConversionGoal(inner, context, typeHint))
    .filter((item): item is ConversionGoal => Boolean(item));
}

export async function getConversionGoalsByIds(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
  conversionGoalIds: string[],
): Promise<ConversionGoal[]> {
  const xml = await campaignSoap(
    context,
    "GetConversionGoalsByIds",
    `${longArrayXml("ConversionGoalIds", conversionGoalIds)}<ConversionGoalTypes>${ALL_CONVERSION_GOAL_TYPES}</ConversionGoalTypes><ReturnAdditionalFields>${CONVERSION_GOAL_ADDITIONAL_FIELDS}</ReturnAdditionalFields>`,
  );
  return conversionGoalBlocks(xml)
    .map(({ inner, typeHint }) => mapConversionGoal(inner, context, typeHint))
    .filter((item): item is ConversionGoal => Boolean(item));
}

export async function getUetTagsByAccount(
  context: MicrosoftRequestContext & { customerId: string; accountId: string },
): Promise<UetTag[]> {
  const xml = await campaignSoap(
    context,
    "GetUetTagsByIds",
    `<TagIds i:nil="true"/>`,
  );
  return childBlocks(xml, "UetTag")
    .map((block) => mapUetTag(block, context))
    .filter((item): item is UetTag => Boolean(item));
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
