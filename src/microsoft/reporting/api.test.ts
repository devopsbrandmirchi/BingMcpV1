import { describe, expect, it } from "vitest";
import { zipSync } from "fflate";
import {
  buildReportColumns,
  isDiscardedReportRow,
  mapRow,
  parseCsv,
  rankSearchQueryRows,
  resolveAggregation,
  unzipFirstTextFile,
} from "@/microsoft/reporting/api";

describe("report parsing", () => {
  it("parses Microsoft CSV report rows", () => {
    const csv = [
      "Report Name,Campaign Performance",
      "TimePeriod,CampaignId,CampaignName,Impressions,Clicks,Spend",
      "2026-08-01,123,Campaign A,5000,200,100.25",
      '"© Microsoft"',
    ].join("\n");
    const rows = parseCsv(csv).map(mapRow);
    expect(rows).toHaveLength(1);
    expect(rows[0]?.campaignId).toBe("123");
    expect(rows[0]?.impressions).toBe(5000);
    expect(rows[0]?.spend).toBe(100.25);
  });

  it("maps search query rows separately from keywords", () => {
    const csv = [
      "SearchQuery,Keyword,CampaignName,Impressions,Clicks,Spend,Conversions,ConversionRate,CostPerConversion",
      "newmar motorhome dealer washington,newmar dealer,WA | Search | Newmar,125,18,42.50,1,8.00,42.50",
      ",newmar dealer,WA | Search | Newmar,10,0,0,0,0,",
      "Total,,Totals,135,18,42.50,1,,",
    ].join("\n");
    const parsed = parseCsv(csv).filter((raw) => !isDiscardedReportRow(raw)).map(mapRow);
    expect(parsed).toHaveLength(1);
    expect(parsed[0]?.searchQuery).toBe("newmar motorhome dealer washington");
    expect(parsed[0]?.keyword).toBe("newmar dealer");
    expect(parsed[0]?.impressions).toBe(125);
    expect(parsed[0]?.costPerConversion).toBe(42.5);
  });

  it("computes cost per conversion only when conversions are greater than zero", () => {
    const zero = mapRow({
      SearchQuery: "rv dealer",
      Keyword: "rv",
      Spend: "20",
      Conversions: "0",
      Clicks: "4",
      Impressions: "40",
    });
    expect(zero.searchQuery).toBe("rv dealer");
    expect(zero.keyword).toBe("rv");
    expect(zero.costPerConversion).toBeNull();

    const computed = mapRow({
      SearchQuery: "rv dealer salem",
      Spend: "20",
      Conversions: "2",
    });
    expect(computed.costPerConversion).toBe(10);
  });

  it("returns no rows for an empty search query report", () => {
    expect(parseCsv("")).toEqual([]);
    expect(parseCsv("SearchQuery,Impressions\n").filter((raw) => !isDiscardedReportRow(raw))).toEqual([]);
  });

  it("sorts search queries and applies top N and conversion filters", () => {
    const rows = [
      { searchQuery: "a", spend: 10, clicks: 2, impressions: 20, conversions: 1 },
      { searchQuery: "b", spend: 50, clicks: 8, impressions: 80, conversions: 0 },
      { searchQuery: "c", spend: 5, clicks: 0, impressions: 12, conversions: 0 },
    ];
    expect(rankSearchQueryRows(rows, { sortBy: "spend", limit: 1 })[0]?.searchQuery).toBe("b");
    expect(rankSearchQueryRows(rows, { sortBy: "clicks", limit: 2 }).map((row) => row.searchQuery)).toEqual(["b", "a"]);
    expect(
      rankSearchQueryRows(rows, { sortBy: "spend", minClicks: 1, maxConversions: 0 }).map((row) => row.searchQuery),
    ).toEqual(["b"]);
  });

  it("maps revenue columns and leaves them undefined when Microsoft returns blanks", () => {
    const withRevenue = mapRow({
      TimePeriod: "2026-08-17",
      Spend: "20.97",
      Conversions: "7",
      Revenue: "1,250.50",
      ReturnOnAdSpend: "59.63",
      RevenuePerConversion: "178.64",
    });
    expect(withRevenue.revenue).toBe(1250.5);
    expect(withRevenue.returnOnAdSpend).toBe(59.63);
    expect(withRevenue.revenuePerConversion).toBe(178.64);

    const blank = mapRow({ TimePeriod: "2026-08-18", Spend: "5", Conversions: "0", Revenue: "", ReturnOnAdSpend: "" });
    expect(blank.revenue).toBeUndefined();
    expect(blank.returnOnAdSpend).toBeUndefined();
    expect(JSON.parse(JSON.stringify(blank))).not.toHaveProperty("revenue");
  });

  it("requests revenue columns for account, campaign, ad group and keyword reports only", () => {
    for (const type of [
      "AccountPerformanceReportRequest",
      "CampaignPerformanceReportRequest",
      "AdGroupPerformanceReportRequest",
      "KeywordPerformanceReportRequest",
    ] as const) {
      const columns = buildReportColumns(type, "Daily");
      expect(columns).toContain("Revenue");
      expect(columns).toContain("ReturnOnAdSpend");
      expect(columns).toContain("RevenuePerConversion");
      expect(columns).toContain("TimePeriod");
    }
    const searchQuery = buildReportColumns("SearchQueryPerformanceReportRequest", "Summary");
    expect(searchQuery).not.toContain("Revenue");
  });

  it("keeps TimePeriod for Monthly and drops it for Summary", () => {
    expect(buildReportColumns("CampaignPerformanceReportRequest", "Monthly")).toContain("TimePeriod");
    expect(buildReportColumns("CampaignPerformanceReportRequest", "Summary")).not.toContain("TimePeriod");
  });

  it("defaults to Daily and always forces Summary for search queries", () => {
    expect(resolveAggregation("AccountPerformanceReportRequest")).toBe("Daily");
    expect(resolveAggregation("AccountPerformanceReportRequest", "Monthly")).toBe("Monthly");
    expect(resolveAggregation("SearchQueryPerformanceReportRequest", "Daily")).toBe("Summary");
  });

  it("unzips the first text file from a report archive", () => {
    const archive = zipSync({
      "report.csv": new TextEncoder().encode("TimePeriod,Impressions\n2026-08-01,10\n"),
    });
    const text = unzipFirstTextFile(archive);
    expect(text).toContain("Impressions");
  });
});
