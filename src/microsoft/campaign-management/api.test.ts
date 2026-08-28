import { beforeEach, describe, expect, it, vi } from "vitest";
import { getCampaignsByAccountId, getAdGroupsByCampaignId } from "@/microsoft/campaign-management/api";
import { setRequiredEnv } from "@/test/env";

describe("Campaign Management adapters", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("maps campaigns and ad groups", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          JSON.stringify({
            Campaigns: [
              { Id: 55, Name: "Search", Status: "Active", CampaignType: "Search", DailyBudget: 20 },
            ],
          }),
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          JSON.stringify({
            AdGroups: [{ Id: 77, Name: "Brand", Status: "Paused" }],
          }),
      });
    vi.stubGlobal("fetch", fetchMock);

    const context = { accessToken: "token", customerId: "11", accountId: "123" };
    const campaigns = await getCampaignsByAccountId(context);
    expect(campaigns[0]?.campaignId).toBe("55");
    expect(campaigns[0]?.name).toBe("Search");
    expect(JSON.parse(String(fetchMock.mock.calls[0]?.[1]?.body))).toEqual({
      AccountId: 123,
      CampaignType: "Search Shopping DynamicSearchAds Audience Hotel PerformanceMax App",
    });

    const adGroups = await getAdGroupsByCampaignId(context, "55");
    expect(adGroups[0]?.adGroupId).toBe("77");
    expect(adGroups[0]?.status).toBe("Paused");
    expect(JSON.parse(String(fetchMock.mock.calls[1]?.[1]?.body))).toEqual({
      CampaignId: 55,
    });
  });
});
