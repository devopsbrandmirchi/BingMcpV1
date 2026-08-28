import { beforeEach, describe, expect, it, vi } from "vitest";
import { getCampaignsByAccountId, getAdGroupsByCampaignId } from "@/microsoft/campaign-management/api";
import { setRequiredEnv } from "@/test/env";

describe("Campaign Management adapters", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("maps SOAP campaigns and ad groups", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetCampaignsByAccountIdResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><Campaigns><Campaign><Id>55</Id><Name>Search</Name><Status>Active</Status><CampaignType>Search</CampaignType><DailyBudget>20</DailyBudget></Campaign></Campaigns></GetCampaignsByAccountIdResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetAdGroupsByCampaignIdResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><AdGroups><AdGroup><Id>77</Id><Name>Brand</Name><Status>Paused</Status></AdGroup></AdGroups></GetAdGroupsByCampaignIdResponse></s:Body></s:Envelope>`,
      });
    vi.stubGlobal("fetch", fetchMock);

    const context = { accessToken: "token", customerId: "11", accountId: "123" };
    const campaigns = await getCampaignsByAccountId(context);
    expect(campaigns[0]?.campaignId).toBe("55");
    expect(campaigns[0]?.name).toBe("Search");
    expect(String(fetchMock.mock.calls[0]?.[0])).toContain("CampaignManagementService.svc");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("GetCampaignsByAccountIdRequest");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("<CampaignType>Search Shopping DynamicSearchAds Audience Hotel PerformanceMax App</CampaignType>");
    expect(String(fetchMock.mock.calls[0]?.[1]?.headers?.SOAPAction)).toBe('"GetCampaignsByAccountId"');

    const adGroups = await getAdGroupsByCampaignId(context, "55");
    expect(adGroups[0]?.adGroupId).toBe("77");
    expect(adGroups[0]?.status).toBe("Paused");
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).toContain("<CampaignId>55</CampaignId>");
  });
});
