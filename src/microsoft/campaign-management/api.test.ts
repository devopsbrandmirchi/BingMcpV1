import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  getAdGroupsByCampaignId,
  getCampaignsByAccountId,
  getConversionGoalsByAccount,
  getConversionGoalsByIds,
  getUetTagsByAccount,
} from "@/microsoft/campaign-management/api";
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

  it("maps SOAP conversion goals and UET tags", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetConversionGoalsByIdsResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><ConversionGoals><ConversionGoal i:type="UrlGoal"><Id>901</Id><Name>Thank you page</Name><Status>Active</Status><TagId>44001</TagId><TrackingStatus>TagActive</TrackingStatus><Scope>Account</Scope><CountType>Unique</CountType><ExcludeFromBidding>false</ExcludeFromBidding><GoalCategory>Purchase</GoalCategory><ConversionWindowInMinutes>43200</ConversionWindowInMinutes><Revenue><Type>FixedValue</Type><Value>25</Value><CurrencyCode>USD</CurrencyCode></Revenue><UrlExpression>/thank-you</UrlExpression><UrlOperator>Contains</UrlOperator></ConversionGoal></ConversionGoals></GetConversionGoalsByIdsResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetConversionGoalsByIdsResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><ConversionGoals><ConversionGoal i:type="EventGoal"><Id>902</Id><Name>Lead form</Name><Type>Event</Type><Status>Active</Status><TagId>44001</TagId><ActionExpression>submit</ActionExpression></ConversionGoal></ConversionGoals></GetConversionGoalsByIdsResponse></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetUetTagsByIdsResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><UetTags><UetTag><Id>44001</Id><Name>Zoomers site</Name><Description>Main site tag</Description><TrackingStatus>Active</TrackingStatus><Industry>Automotive</Industry><TrackingScript>do-not-return</TrackingScript></UetTag></UetTags></GetUetTagsByIdsResponse></s:Body></s:Envelope>`,
      });
    vi.stubGlobal("fetch", fetchMock);

    const context = { accessToken: "token", customerId: "11", accountId: "188405633" };
    const goals = await getConversionGoalsByAccount(context);
    expect(goals[0]?.conversionGoalId).toBe("901");
    expect(goals[0]?.name).toBe("Thank you page");
    expect(goals[0]?.type).toBe("Url");
    expect(goals[0]?.tagId).toBe("44001");
    expect(goals[0]?.trackingStatus).toBe("TagActive");
    expect(goals[0]?.urlExpression).toBe("/thank-you");
    expect(goals[0]?.revenue?.value).toBe(25);
    expect(goals[0]?.revenue?.currencyCode).toBe("USD");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("GetConversionGoalsByIdsRequest");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain('<ConversionGoalIds i:nil="true"/>');
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("Url Duration PagesViewedPerVisit Event");
    expect(String(fetchMock.mock.calls[0]?.[1]?.headers?.SOAPAction)).toBe('"GetConversionGoalsByIds"');

    const one = await getConversionGoalsByIds(context, ["902"]);
    expect(one[0]?.conversionGoalId).toBe("902");
    expect(one[0]?.type).toBe("Event");
    expect(one[0]?.actionExpression).toBe("submit");
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).toContain("<a1:long>902</a1:long>");

    const tags = await getUetTagsByAccount(context);
    expect(tags[0]?.uetTagId).toBe("44001");
    expect(tags[0]?.name).toBe("Zoomers site");
    expect(tags[0]?.trackingStatus).toBe("Active");
    expect(JSON.stringify(tags[0])).not.toContain("do-not-return");
    expect(String(fetchMock.mock.calls[2]?.[1]?.body)).toContain("GetUetTagsByIdsRequest");
    expect(String(fetchMock.mock.calls[2]?.[1]?.body)).toContain('<TagIds i:nil="true"/>');
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).not.toContain("InStoreTransaction");
  });

  it("retries conversion goal list without pilot-gated types", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce({
        ok: false,
        status: 500,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><s:Fault><faultstring>The customer is not enabled for this pilot.</faultstring><detail><ApiFaultDetail xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><TrackingId>track-pilot</TrackingId><OperationErrors><OperationError><Code>0</Code><ErrorCode>AppDownloadPilotNotEnabledForCustomer</ErrorCode><Message>The customer is not enabled for this pilot.</Message></OperationError></OperationErrors></ApiFaultDetail></detail></s:Fault></s:Body></s:Envelope>`,
      })
      .mockResolvedValueOnce({
        ok: true,
        headers: new Headers(),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><GetConversionGoalsByIdsResponse xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><ConversionGoals><ConversionGoal i:type="UrlGoal"><Id>901</Id><Name>Thank you page</Name><Type>Url</Type></ConversionGoal></ConversionGoals></GetConversionGoalsByIdsResponse></s:Body></s:Envelope>`,
      });
    vi.stubGlobal("fetch", fetchMock);

    const goals = await getConversionGoalsByAccount({
      accessToken: "token",
      customerId: "11",
      accountId: "188405633",
    });
    expect(goals[0]?.conversionGoalId).toBe("901");
    expect(String(fetchMock.mock.calls[0]?.[1]?.body)).toContain("AppDownload");
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).not.toContain("AppDownload");
    expect(String(fetchMock.mock.calls[1]?.[1]?.body)).toContain("Url Duration PagesViewedPerVisit Event");
  });
});
