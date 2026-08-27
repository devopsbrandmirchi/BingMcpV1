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
  impressions?: number;
  clicks?: number;
  spend?: number;
  ctr?: number;
  averageCpc?: number;
  conversions?: number;
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
