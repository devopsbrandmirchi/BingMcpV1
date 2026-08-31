import { unzipSync } from "fflate";
import { assertReportWindow, toMicrosoftReportDate, type ReportTimePeriod } from "@/lib/dates";
import { ReportError } from "@/lib/errors";
import { logger } from "@/lib/logger";
import { peekOperatorContext } from "@/lib/request-context";
import { microsoftSoapRequest } from "@/microsoft/client/soap";
import { childText, longArrayXml, stripXmlNamespaces, xmlEscape } from "@/microsoft/client/xml";
import { REPORTING_NAMESPACE, REPORTING_SOAP_URL } from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { NormalizedReport, ReportRow } from "@/microsoft/models/types";

const POLL_INTERVAL_MS = 2000;
const MAX_POLL_MS = 240_000;
const MAX_REPORT_ROWS = 5000;

export type PerformanceReportType =
  | "AccountPerformanceReportRequest"
  | "CampaignPerformanceReportRequest"
  | "AdGroupPerformanceReportRequest"
  | "KeywordPerformanceReportRequest"
  | "SearchQueryPerformanceReportRequest";

export const SEARCH_QUERY_SORT_FIELDS = ["spend", "clicks", "impressions", "conversions"] as const;
export type SearchQuerySortField = (typeof SEARCH_QUERY_SORT_FIELDS)[number];

const COLUMNS_BY_TYPE: Record<PerformanceReportType, string[]> = {
  AccountPerformanceReportRequest: [
    "TimePeriod",
    "AccountId",
    "AccountName",
    "Impressions",
    "Clicks",
    "Spend",
    "Ctr",
    "AverageCpc",
    "Conversions",
  ],
  CampaignPerformanceReportRequest: [
    "TimePeriod",
    "AccountId",
    "AccountName",
    "CampaignId",
    "CampaignName",
    "Impressions",
    "Clicks",
    "Spend",
    "Ctr",
    "AverageCpc",
    "Conversions",
  ],
  AdGroupPerformanceReportRequest: [
    "TimePeriod",
    "AccountId",
    "CampaignId",
    "CampaignName",
    "AdGroupId",
    "AdGroupName",
    "Impressions",
    "Clicks",
    "Spend",
    "Ctr",
    "AverageCpc",
    "Conversions",
  ],
  KeywordPerformanceReportRequest: [
    "TimePeriod",
    "AccountId",
    "CampaignId",
    "CampaignName",
    "AdGroupId",
    "AdGroupName",
    "KeywordId",
    "Keyword",
    "Impressions",
    "Clicks",
    "Spend",
    "Ctr",
    "AverageCpc",
    "Conversions",
  ],
  SearchQueryPerformanceReportRequest: [
    "AccountId",
    "CampaignId",
    "CampaignName",
    "AdGroupId",
    "AdGroupName",
    "SearchQuery",
    "Keyword",
    "Impressions",
    "Clicks",
    "Spend",
    "Ctr",
    "AverageCpc",
    "Conversions",
    "ConversionRate",
    "CostPerConversion",
  ],
};

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseNumber(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const cleaned = value.replace(/[$,%]/g, "").trim();
  if (!cleaned) {
    return undefined;
  }
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function parseCsv(text: string): Array<Record<string, string>> {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((line) => line.length > 0);
  const headerIndex = lines.findIndex((line) =>
    /TimePeriod|GregorianDate|AccountId|Impressions|SearchQuery/i.test(line),
  );
  if (headerIndex < 0) {
    return [];
  }
  const headers = splitCsvLine(lines[headerIndex] ?? "");
  const rows: Array<Record<string, string>> = [];
  for (const line of lines.slice(headerIndex + 1)) {
    if (/copyright|©|microsoft advertising/i.test(line) && !/^\d{4}-\d{2}-\d{2}/.test(line)) {
      break;
    }
    const cells = splitCsvLine(line);
    if (cells.length === 0 || cells.every((cell) => !cell)) {
      continue;
    }
    const row: Record<string, string> = {};
    headers.forEach((header, index) => {
      row[header] = cells[index] ?? "";
    });
    rows.push(row);
  }
  return rows;
}

function splitCsvLine(line: string): string[] {
  const cells: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if ((char === "," || char === "\t") && !inQuotes) {
      cells.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  cells.push(current.trim());
  return cells;
}

async function downloadReportArchive(url: string): Promise<Uint8Array> {
  const retryable = new Set([409, 404, 429, 500, 502, 503]);
  let lastStatus = 0;
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const response = await fetch(url, { redirect: "follow" });
    lastStatus = response.status;
    if (response.ok) {
      return new Uint8Array(await response.arrayBuffer());
    }
    if (!retryable.has(response.status) || attempt === 4) {
      throw new ReportError(`The report download failed with HTTP ${response.status}.`);
    }
    logger.warn("Microsoft report download retry", { status: response.status, attempt });
    await sleep(1000 * 2 ** attempt);
  }
  throw new ReportError(`The report download failed with HTTP ${lastStatus}.`);
}

function unzipFirstTextFile(buffer: Uint8Array): string {
  const files = unzipSync(buffer);
  const names = Object.keys(files);
  const preferred = names.find((name) => /\.(csv|tsv|txt)$/i.test(name)) ?? names[0];
  if (!preferred || !files[preferred]) {
    throw new ReportError("The downloaded Microsoft report archive did not contain a data file.");
  }
  return new TextDecoder("utf-8").decode(files[preferred]);
}

function isDiscardedReportRow(raw: Record<string, string>): boolean {
  const time = (raw.TimePeriod || raw.GregorianDate || "").trim();
  if (/^(total|totals|summary)$/i.test(time)) {
    return true;
  }
  if ("SearchQuery" in raw) {
    const query = (raw.SearchQuery || "").trim();
    if (!query || /^(total|totals|summary)$/i.test(query)) {
      return true;
    }
  }
  return false;
}

function costPerConversion(
  raw: Record<string, string>,
  spend: number | undefined,
  conversions: number | undefined,
): number | null | undefined {
  const fromMicrosoft = parseNumber(raw.CostPerConversion);
  if (fromMicrosoft !== undefined) {
    return fromMicrosoft;
  }
  if (!("CostPerConversion" in raw) && !("SearchQuery" in raw)) {
    return undefined;
  }
  if (conversions && conversions > 0 && spend !== undefined) {
    return spend / conversions;
  }
  return null;
}

function mapRow(raw: Record<string, string>): ReportRow {
  const spend = parseNumber(raw.Spend);
  const conversions = parseNumber(raw.Conversions);
  const searchQuery = raw.SearchQuery?.trim() || undefined;
  return {
    date: raw.TimePeriod || raw.GregorianDate || undefined,
    accountId: raw.AccountId || undefined,
    accountName: raw.AccountName || undefined,
    campaignId: raw.CampaignId || undefined,
    campaignName: raw.CampaignName || undefined,
    adGroupId: raw.AdGroupId || undefined,
    adGroupName: raw.AdGroupName || undefined,
    keywordId: raw.KeywordId || undefined,
    keyword: raw.Keyword || undefined,
    searchQuery,
    impressions: parseNumber(raw.Impressions),
    clicks: parseNumber(raw.Clicks),
    spend,
    ctr: parseNumber(raw.Ctr),
    averageCpc: parseNumber(raw.AverageCpc),
    conversions,
    conversionRate: parseNumber(raw.ConversionRate),
    costPerConversion: costPerConversion(raw, spend, conversions),
  };
}

export function rankSearchQueryRows(
  rows: ReportRow[],
  options: {
    sortBy?: SearchQuerySortField;
    limit?: number;
    minClicks?: number;
    maxConversions?: number;
  } = {},
): ReportRow[] {
  let filtered = rows;
  if (options.minClicks !== undefined) {
    filtered = filtered.filter((row) => (row.clicks ?? 0) >= options.minClicks!);
  }
  if (options.maxConversions !== undefined) {
    filtered = filtered.filter((row) => (row.conversions ?? 0) <= options.maxConversions!);
  }
  const field = options.sortBy ?? "spend";
  const sorted = [...filtered].sort((left, right) => (right[field] ?? 0) - (left[field] ?? 0));
  const limit = Math.min(Math.max(options.limit ?? 20, 1), 1000);
  return sorted.slice(0, limit);
}

function reportDateXml(isoDate: string): string {
  const { Year, Month, Day } = toMicrosoftReportDate(isoDate);
  return `<Day>${Day}</Day><Month>${Month}</Month><Year>${Year}</Year>`;
}

function buildReportTimeXml(window: { period?: ReportTimePeriod; startDate?: string; endDate?: string }): string {
  if (window.period) {
    return `<Time><PredefinedTime>${xmlEscape(window.period)}</PredefinedTime></Time>`;
  }
  return [
    `<Time>`,
    `<CustomDateRangeEnd>${reportDateXml(window.endDate as string)}</CustomDateRangeEnd>`,
    `<CustomDateRangeStart>${reportDateXml(window.startDate as string)}</CustomDateRangeStart>`,
    `</Time>`,
  ].join("");
}

export async function runPerformanceReport(params: {
  context: MicrosoftRequestContext & { customerId: string; accountId: string };
  type: PerformanceReportType;
  startDate?: string;
  endDate?: string;
  period?: ReportTimePeriod;
  campaignIds?: string[];
  adGroupIds?: string[];
  keywordIds?: string[];
}): Promise<NormalizedReport> {
  const window = assertReportWindow({
    period: params.period,
    startDate: params.startDate,
    endDate: params.endDate,
  });
  const requestId = peekOperatorContext()?.requestId;
  const columns = COLUMNS_BY_TYPE[params.type];
  const columnTag = params.type.replace("Request", "Column");

  const campaignScope = params.campaignIds?.length
    ? `<Campaigns>${params.campaignIds
        .map(
          (campaignId) =>
            `<CampaignReportScope><AccountId>${xmlEscape(params.context.accountId)}</AccountId><CampaignId>${xmlEscape(campaignId)}</CampaignId></CampaignReportScope>`,
        )
        .join("")}</Campaigns>`
    : "";
  const adGroupScope = params.adGroupIds?.length
    ? `<AdGroups>${params.adGroupIds
        .map(
          (adGroupId) =>
            `<AdGroupReportScope><AccountId>${xmlEscape(params.context.accountId)}</AccountId>${
              params.campaignIds?.[0] ? `<CampaignId>${xmlEscape(params.campaignIds[0])}</CampaignId>` : ""
            }<AdGroupId>${xmlEscape(adGroupId)}</AdGroupId></AdGroupReportScope>`,
        )
        .join("")}</AdGroups>`
    : "";

  const submitXml = await microsoftSoapRequest({
    service: "reporting",
    url: REPORTING_SOAP_URL,
    action: "SubmitGenerateReport",
    namespace: REPORTING_NAMESPACE,
    context: params.context,
    bodyXml: [
      `<SubmitGenerateReportRequest xmlns="${REPORTING_NAMESPACE}">`,
      `<ReportRequest i:type="a:${params.type}" xmlns:a="${REPORTING_NAMESPACE}">`,
      `<ExcludeColumnHeaders>false</ExcludeColumnHeaders>`,
      `<ExcludeReportFooter>true</ExcludeReportFooter>`,
      `<ExcludeReportHeader>true</ExcludeReportHeader>`,
      `<Format>Csv</Format>`,
      `<FormatVersion>2.0</FormatVersion>`,
      `<ReportName>${xmlEscape(`bing-mcp-${params.type}`)}</ReportName>`,
      `<ReturnOnlyCompleteData>false</ReturnOnlyCompleteData>`,
      `<Aggregation>${params.type === "SearchQueryPerformanceReportRequest" ? "Summary" : "Daily"}</Aggregation>`,
      `<Columns>${columns.map((column) => `<${columnTag}>${xmlEscape(column)}</${columnTag}>`).join("")}</Columns>`,
      `<Scope>${longArrayXml("AccountIds", [params.context.accountId])}${campaignScope}${adGroupScope}</Scope>`,
      buildReportTimeXml(window),
      `</ReportRequest>`,
      `</SubmitGenerateReportRequest>`,
    ].join(""),
  });

  const reportRequestId = childText(stripXmlNamespaces(submitXml), "ReportRequestId");
  if (!reportRequestId) {
    throw new ReportError("Microsoft Advertising did not return a report request ID.");
  }

  const started = Date.now();
  let downloadUrl: string | undefined;
  while (Date.now() - started < MAX_POLL_MS) {
    const polledXml = await microsoftSoapRequest({
      service: "reporting",
      url: REPORTING_SOAP_URL,
      action: "PollGenerateReport",
      namespace: REPORTING_NAMESPACE,
      context: params.context,
      bodyXml: `<PollGenerateReportRequest xmlns="${REPORTING_NAMESPACE}"><ReportRequestId>${xmlEscape(reportRequestId)}</ReportRequestId></PollGenerateReportRequest>`,
    });
    const polled = stripXmlNamespaces(polledXml);
    const status = childText(polled, "Status");
    logger.info("Microsoft report poll", { requestId, reportRequestId, status });
    if (status === "Success") {
      downloadUrl = childText(polled, "ReportDownloadUrl") ?? undefined;
      if (downloadUrl) {
        break;
      }
    }
    if (status === "Error" || status === "Failed") {
      throw new ReportError("Microsoft Advertising failed to generate the report.");
    }
    await sleep(POLL_INTERVAL_MS);
  }

  if (!downloadUrl) {
    throw new ReportError("The Microsoft Advertising report timed out before completion.", "report_timeout");
  }

  const bytes = await downloadReportArchive(downloadUrl);
  const csv = unzipFirstTextFile(bytes);
  const parsed = parseCsv(csv).filter((raw) => !isDiscardedReportRow(raw)).map(mapRow);
  const truncated = parsed.length > MAX_REPORT_ROWS;
  const rows = truncated ? parsed.slice(0, MAX_REPORT_ROWS) : parsed;

  const summary = rows.reduce(
    (acc, row) => ({
      impressions: acc.impressions + (row.impressions ?? 0),
      clicks: acc.clicks + (row.clicks ?? 0),
      spend: acc.spend + (row.spend ?? 0),
      conversions: acc.conversions + (row.conversions ?? 0),
    }),
    { impressions: 0, clicks: 0, spend: 0, conversions: 0 } as {
      impressions: number;
      clicks: number;
      spend: number;
      conversions: number;
    },
  );

  return {
    summary,
    rows,
    truncated,
    rowCount: rows.length,
  };
}

export { parseCsv, mapRow, unzipFirstTextFile, isDiscardedReportRow };
