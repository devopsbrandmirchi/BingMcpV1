export interface MicrosoftUser {
  id: string | null;
  customerId: string | null;
  userName: string | null;
  email: string | null;
  firstName: string | null;
  lastName: string | null;
  status: string | null;
}

export interface MicrosoftCustomerRole {
  customerId: string;
  accountIds: string[];
  linkedAccountIds: string[];
  roleId: number | null;
}

export interface Campaign {
  campaignId: string;
  accountId: string;
  customerId: string;
  name: string | null;
  status: string | null;
  campaignType: string | null;
  budget: number | null;
  budgetType: string | null;
  timeZone: string | null;
}

export interface AdGroup {
  adGroupId: string;
  campaignId: string;
  accountId: string;
  customerId: string;
  name: string | null;
  status: string | null;
}

export interface Ad {
  adId: string;
  adGroupId: string;
  campaignId?: string;
  accountId: string;
  customerId: string;
  type: string | null;
  status: string | null;
  editorialStatus: string | null;
  headline?: string | null;
  text?: string | null;
  displayUrl?: string | null;
  finalUrls?: string[];
}

export interface ConversionGoalRevenue {
  type: string | null;
  value: number | null;
  currencyCode: string | null;
}

export interface ConversionGoal {
  conversionGoalId: string;
  accountId: string;
  customerId: string;
  name: string | null;
  type: string | null;
  status: string | null;
  trackingStatus: string | null;
  tagId: string | null;
  scope: string | null;
  countType: string | null;
  goalCategory: string | null;
  excludeFromBidding: boolean | null;
  isAutoGoal: boolean | null;
  isEnhancedConversionsEnabled: boolean | null;
  conversionWindowInMinutes: number | null;
  viewThroughConversionWindowInMinutes: number | null;
  attributionModelType: string | null;
  revenue: ConversionGoalRevenue | null;
  urlExpression: string | null;
  urlOperator: string | null;
  actionExpression: string | null;
  categoryExpression: string | null;
  labelExpression: string | null;
  minimumDurationInSeconds: number | null;
  minimumPagesViewed: number | null;
}

export interface UetTag {
  uetTagId: string;
  accountId: string;
  customerId: string;
  name: string | null;
  description: string | null;
  trackingStatus: string | null;
  industry: string | null;
}

export interface Keyword {
  keywordId: string;
  adGroupId: string;
  accountId: string;
  customerId: string;
  text: string | null;
  matchType: string | null;
  status: string | null;
  bid?: number | null;
}

export interface ReportSummary {
  impressions: number;
  clicks: number;
  spend: number;
  conversions?: number;
}

export interface ReportRow {
  date?: string;
  accountId?: string;
  accountName?: string;
  campaignId?: string;
  campaignName?: string;
  adGroupId?: string;
  adGroupName?: string;
  keywordId?: string;
  keyword?: string;
  searchQuery?: string;
  impressions?: number;
  clicks?: number;
  spend?: number;
  ctr?: number;
  averageCpc?: number;
  conversions?: number;
  conversionRate?: number;
  costPerConversion?: number | null;
}

export interface NormalizedReport {
  summary: ReportSummary;
  rows: ReportRow[];
  truncated: boolean;
  rowCount: number;
}

export interface PagedResult<T> {
  items: T[];
  total: number;
  limit: number;
  offset: number;
}
