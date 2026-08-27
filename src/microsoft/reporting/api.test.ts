import { describe, expect, it } from "vitest";
import { zipSync } from "fflate";
import { mapRow, parseCsv, unzipFirstTextFile } from "@/microsoft/reporting/api";

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

  it("unzips the first text file from a report archive", () => {
    const archive = zipSync({
      "report.csv": new TextEncoder().encode("TimePeriod,Impressions\n2026-08-01,10\n"),
    });
    const text = unzipFirstTextFile(archive);
    expect(text).toContain("Impressions");
  });
});
