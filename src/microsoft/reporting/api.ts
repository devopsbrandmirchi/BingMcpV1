import { unzipSync } from "fflate";
import { assertDateRange, toMicrosoftReportDate } from "@/lib/dates";
import { ReportError } from "@/lib/errors";
import { logger } from "@/lib/logger";
import { peekOperatorContext } from "@/lib/request-context";
import { microsoftJsonRequest } from "@/microsoft/client/http";
import { toMicrosoftLong, toMicrosoftLongs } from "@/microsoft/client/ids";
import { REPORTING_BASE } from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { NormalizedReport, ReportRow } from "@/microsoft/models/types";

const POLL_INTERVAL_MS = 2000;
const MAX_POLL_MS = 240_000;
const MAX_REPORT_ROWS = 5000;

export type PerformanceReportType =
  | "AccountPerformanceReportRequest"
  | "CampaignPerformanceReportRequest"
  | "AdGroupPerformanceReportRequest"
  | "KeywordPerformanceReportRequest";

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
  const headerIndex = lines.findIndex((line) => /TimePeriod|GregorianDate|AccountId|Impressions/i.test(line));
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

function unzipFirstTextFile(buffer: Uint8Array): string {
  const files = unzipSync(buffer);
  const names = Object.keys(files);
  const preferred = names.find((name) => /\.(csv|tsv|txt)$/i.test(name)) ?? names[0];
  if (!preferred || !files[preferred]) {
    throw new ReportError("The downloaded Microsoft report archive did not contain a data file.");
  }
  return new TextDecoder("utf-8").decode(files[preferred]);
}

function mapRow(raw: Record<string, string>): ReportRow {
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
    impressions: parseNumber(raw.Impressions),
    clicks: parseNumber(raw.Clicks),
    spend: parseNumber(raw.Spend),
    ctr: parseNumber(raw.Ctr),
    averageCpc: parseNumber(raw.AverageCpc),
    conversions: parseNumber(raw.Conversions),
  };
}

export async function runPerformanceReport(params: {
  context: MicrosoftRequestContext & { customerId: string; accountId: string };
  type: PerformanceReportType;
  startDate: string;
  endDate: string;
  campaignIds?: string[];
  adGroupIds?: string[];
  keywordIds?: string[];
}): Promise<NormalizedReport> {
  const range = assertDateRange(params.startDate, params.endDate);
  const requestId = peekOperatorContext()?.requestId;
  const columns = COLUMNS_BY_TYPE[params.type];

  const accountId = toMicrosoftLong(params.context.accountId, "accountId");
  const campaignIds = params.campaignIds?.length
    ? toMicrosoftLongs(params.campaignIds, "campaignId")
    : undefined;
  const adGroupIds = params.adGroupIds?.length
    ? toMicrosoftLongs(params.adGroupIds, "adGroupId")
    : undefined;

  const scope: Record<string, unknown> = {
    AccountIds: [accountId],
  };
  if (campaignIds) {
    scope.Campaigns = campaignIds.map((campaignId) => ({
      AccountId: accountId,
      CampaignId: campaignId,
    }));
  }
  if (adGroupIds) {
    scope.AdGroups = adGroupIds.map((adGroupId) => ({
      AccountId: accountId,
      ...(campaignIds?.[0] ? { CampaignId: campaignIds[0] } : {}),
      AdGroupId: adGroupId,
    }));
  }

  const reportRequest: Record<string, unknown> = {
    Type: params.type,
    Format: "Csv",
    FormatVersion: "2.0",
    ReportName: `bing-mcp-${params.type}`,
    ReturnOnlyCompleteData: false,
    Aggregation: "Daily",
    ExcludeColumnHeaders: false,
    ExcludeReportFooter: true,
    ExcludeReportHeader: true,
    Columns: columns,
    Scope: scope,
    Time: {
      CustomDateRangeStart: toMicrosoftReportDate(range.startDate),
      CustomDateRangeEnd: toMicrosoftReportDate(range.endDate),
    },
  };

  const submitted = await microsoftJsonRequest<{ ReportRequestId?: string }>({
    service: "reporting",
    url: `${REPORTING_BASE}/GenerateReport/Submit`,
    context: params.context,
    body: { ReportRequest: reportRequest },
  });

  const reportRequestId = submitted.ReportRequestId;
  if (!reportRequestId) {
    throw new ReportError("Microsoft Advertising did not return a report request ID.");
  }

  const started = Date.now();
  let downloadUrl: string | undefined;
  while (Date.now() - started < MAX_POLL_MS) {
    const polled = await microsoftJsonRequest<{
      ReportRequestStatus?: { Status?: string; ReportDownloadUrl?: string };
    }>({
      service: "reporting",
      url: `${REPORTING_BASE}/GenerateReport/Poll`,
      context: params.context,
      body: { ReportRequestId: reportRequestId },
    });
    const status = polled.ReportRequestStatus?.Status;
    logger.info("Microsoft report poll", { requestId, reportRequestId, status });
    if (status === "Success") {
      downloadUrl = polled.ReportRequestStatus?.ReportDownloadUrl;
      break;
    }
    if (status === "Error" || status === "Failed") {
      throw new ReportError("Microsoft Advertising failed to generate the report.");
    }
    await sleep(POLL_INTERVAL_MS);
  }

  if (!downloadUrl) {
    throw new ReportError("The Microsoft Advertising report timed out before completion.", "report_timeout");
  }

  const download = await fetch(downloadUrl);
  if (!download.ok) {
    throw new ReportError(`The report download failed with HTTP ${download.status}.`);
  }
  const bytes = new Uint8Array(await download.arrayBuffer());
  const csv = unzipFirstTextFile(bytes);
  const parsed = parseCsv(csv).map(mapRow);
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

export { parseCsv, mapRow, unzipFirstTextFile };
