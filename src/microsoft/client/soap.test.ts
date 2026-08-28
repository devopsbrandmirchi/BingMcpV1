import { beforeEach, describe, expect, it, vi } from "vitest";
import { buildSoapEnvelope, microsoftSoapRequest } from "@/microsoft/client/soap";
import { setRequiredEnv } from "@/test/env";

describe("Microsoft SOAP client", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("builds a SOAP envelope with AuthenticationToken and CustomerId", () => {
    const envelope = buildSoapEnvelope({
      action: "GetCampaignsByAccountId",
      namespace: "https://bingads.microsoft.com/CampaignManagement/v13",
      context: { accessToken: "token", customerId: "11", accountId: "123" },
      bodyXml: "<GetCampaignsByAccountIdRequest></GetCampaignsByAccountIdRequest>",
    });
    expect(envelope).toContain("<AuthenticationToken>token</AuthenticationToken>");
    expect(envelope).toContain("<CustomerId>11</CustomerId>");
    expect(envelope).toContain("<CustomerAccountId>123</CustomerAccountId>");
    expect(envelope).toContain("<DeveloperToken>test-ads-developer-token</DeveloperToken>");
  });

  it("surfaces SOAP fault OperationErrors", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: false,
        status: 500,
        headers: new Headers({ TrackingId: "track-1" }),
        text: async () =>
          `<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><s:Fault><faultstring>The request message is null.</faultstring><detail><ApiFaultDetail xmlns="https://bingads.microsoft.com/CampaignManagement/v13"><TrackingId>track-1</TrackingId><OperationErrors><OperationError><Code>100</Code><ErrorCode>NullRequest</ErrorCode><Message>The request message is null.</Message></OperationError></OperationErrors></ApiFaultDetail></detail></s:Fault></s:Body></s:Envelope>`,
      }),
    );

    await expect(
      microsoftSoapRequest({
        service: "campaign-management",
        url: "https://campaign.api.bingads.microsoft.com/Api/Advertiser/CampaignManagement/v13/CampaignManagementService.svc",
        action: "GetCampaignsByAccountId",
        namespace: "https://bingads.microsoft.com/CampaignManagement/v13",
        context: { accessToken: "token", customerId: "11", accountId: "123" },
        bodyXml: "<GetCampaignsByAccountIdRequest></GetCampaignsByAccountIdRequest>",
      }),
    ).rejects.toMatchObject({
      message: "NullRequest: The request message is null.",
    });
  });
});
