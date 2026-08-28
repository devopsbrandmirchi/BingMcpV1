import { zipSync } from "fflate";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { runPerformanceReport } from "@/microsoft/reporting/api";
import { setRequiredEnv } from "@/test/env";

describe("reporting workflow", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("submits, polls, downloads, and parses a report", async () => {
    const csv = "TimePeriod,AccountId,Impressions,Clicks,Spend\n2026-08-01,123,100,5,12.5\n";
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () => JSON.stringify({ ReportRequestId: "rep-1" }),
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          JSON.stringify({
            ReportRequestStatus: {
              Status: "Success",
              ReportDownloadUrl: "https://example.com/report.zip",
            },
          }),
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "123" },
      type: "AccountPerformanceReportRequest",
      startDate: "2026-08-01",
      endDate: "2026-08-20",
    });
    expect(report.summary.impressions).toBe(100);
    expect(report.summary.clicks).toBe(5);
    expect(report.rows[0]?.accountId).toBe("123");
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("/GenerateReport/Submit");
    expect(JSON.parse(String(fetchMock.mock.calls[0]?.[1]?.body)).ReportRequest).toMatchObject({
      Type: "AccountPerformanceReportRequest",
      FormatVersion: "2.0",
      Scope: { AccountIds: [123] },
    });
    expect(String(fetchMock.mock.calls[1]?.[0])).toContain("/GenerateReport/Poll");
  });
});
