import { zipSync } from "fflate";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { runPerformanceReport } from "@/microsoft/reporting/api";
import { setRequiredEnv } from "@/test/env";

describe("reporting workflow", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("submits, polls, downloads, and parses a SOAP report", async () => {
    const csv = "TimePeriod,AccountId,Impressions,Clicks,Spend\n2026-08-01,123,100,5,12.5\n";
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-1</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/report.zip</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
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
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("ReportingService.svc");
    const submitBody = String(fetchMock.mock.calls[0]?.[1]?.body);
    expect(submitBody).toContain('i:type="a:AccountPerformanceReportRequest"');
    expect(submitBody).toContain("SubmitGenerateReportRequest");
    expect(submitBody.indexOf("CustomDateRangeEnd")).toBeLessThan(submitBody.indexOf("CustomDateRangeStart"));
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).toContain("PollGenerateReportRequest");
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).toContain("rep-1");
  });

  it("uses PredefinedTime for ThisMonth", async () => {
    const csv = "TimePeriod,AccountId,Impressions,Clicks,Spend\n2026-08-01,123,10,1,2\n";
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-2</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/report.zip</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "188405633" },
      type: "AccountPerformanceReportRequest",
      period: "ThisMonth",
    });
    expect(report.summary.spend).toBe(2);
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("<PredefinedTime>ThisMonth</PredefinedTime>");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).not.toContain("CustomDateRangeStart");
  });

  it("retries a 409 report download", async () => {
    const csv = "TimePeriod,AccountId,Impressions,Clicks,Spend\n2026-08-01,123,10,1,2\n";
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-3</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/report.zip?a=1&amp;b=2</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: false,
        status: 409,
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "188405633" },
      type: "AccountPerformanceReportRequest",
      period: "ThisWeek",
    });
    expect(report.summary.spend).toBe(2);
    expect(String(fetchMock.mock.calls[2]?.[0])).toBe("https://example.com/report.zip?a=1&b=2");
  });

  it("submits a Monthly report with revenue columns and sums revenue into the summary", async () => {
    const csv = [
      "TimePeriod,AccountId,Impressions,Clicks,Spend,Conversions,Revenue,ReturnOnAdSpend,RevenuePerConversion",
      "2026-07-01,123,100,10,50,2,300,6,150",
      "2026-08-01,123,200,20,100,3,,,",
    ].join("\n");
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-m</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/monthly.zip</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "123" },
      type: "AccountPerformanceReportRequest",
      startDate: "2026-07-01",
      endDate: "2026-08-30",
      aggregation: "Monthly",
    });
    const submitBody = String(fetchMock.mock.calls[0]?.[1]?.body);
    expect(submitBody).toContain("<Aggregation>Monthly</Aggregation>");
    expect(submitBody).toContain("<AccountPerformanceReportColumn>TimePeriod</AccountPerformanceReportColumn>");
    expect(submitBody).toContain("<AccountPerformanceReportColumn>Revenue</AccountPerformanceReportColumn>");
    expect(submitBody).toContain("<AccountPerformanceReportColumn>ReturnOnAdSpend</AccountPerformanceReportColumn>");
    expect(report.rowCount).toBe(2);
    expect(report.rows[0]?.revenue).toBe(300);
    expect(report.rows[1]?.revenue).toBeUndefined();
    expect(report.summary.revenue).toBe(300);
    expect(report.summary.returnOnAdSpend).toBe(2);
    expect(report.summary.spend).toBe(150);
  });

  it("omits revenue from the summary when no row reports it, and drops TimePeriod for Summary", async () => {
    const csv = "AccountId,Impressions,Clicks,Spend,Conversions,Revenue\n123,100,5,12.5,0,\n";
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-s</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/summary.zip</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "123" },
      type: "AccountPerformanceReportRequest",
      period: "ThisMonth",
      aggregation: "Summary",
    });
    const submitBody = String(fetchMock.mock.calls[0]?.[1]?.body);
    expect(submitBody).toContain("<Aggregation>Summary</Aggregation>");
    expect(submitBody).not.toContain("<AccountPerformanceReportColumn>TimePeriod</AccountPerformanceReportColumn>");
    expect(report.summary).not.toHaveProperty("revenue");
    expect(report.summary).not.toHaveProperty("returnOnAdSpend");
    expect(report.summary.spend).toBe(12.5);
  });

  it("submits a Summary search query report with campaign scope", async () => {
    const csv = [
      "SearchQuery,Keyword,CampaignId,CampaignName,Impressions,Clicks,Spend,Conversions",
      "newmar washington,newmar dealer,555,WA | Search | Newmar,125,18,42.50,1",
    ].join("\n");
    const archive = zipSync({ "report.csv": new TextEncoder().encode(csv) });
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><SubmitGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestId>rep-sq</ReportRequestId></SubmitGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><PollGenerateReportResponse xmlns="https://bingads.microsoft.com/Reporting/v13"><ReportRequestStatus><Status>Success</Status><ReportDownloadUrl>https://example.com/search.zip</ReportDownloadUrl></ReportRequestStatus></PollGenerateReportResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => archive.buffer,
      });
    vi.stubGlobal("fetch", fetchMock);

    const report = await runPerformanceReport({
      context: { accessToken: "token", customerId: "11", accountId: "188378621" },
      type: "SearchQueryPerformanceReportRequest",
      startDate: "2026-08-01",
      endDate: "2026-08-31",
      campaignIds: ["555"],
    });
    expect(report.rows[0]?.searchQuery).toBe("newmar washington");
    expect(report.rows[0]?.keyword).toBe("newmar dealer");
    expect(report.rows[0]?.costPerConversion).toBe(42.5);
    const submitBody = String(fetchMock.mock.calls[0]?.[1]?.body);
    expect(submitBody).toContain('i:type="a:SearchQueryPerformanceReportRequest"');
    expect(submitBody).toContain("<Aggregation>Summary</Aggregation>");
    expect(submitBody).toContain("<SearchQueryPerformanceReportColumn>SearchQuery</SearchQueryPerformanceReportColumn>");
    expect(submitBody).not.toContain("<SearchQueryPerformanceReportColumn>TimePeriod</SearchQueryPerformanceReportColumn>");
    expect(submitBody).toContain("<CampaignId>555</CampaignId>");
    expect(submitBody.indexOf("CustomDateRangeEnd")).toBeLessThan(submitBody.indexOf("CustomDateRangeStart"));
  });
});
