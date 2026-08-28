import { beforeEach, describe, expect, it, vi } from "vitest";
import { microsoftJsonRequest } from "@/microsoft/client/http";
import { setRequiredEnv } from "@/test/env";

describe("Microsoft JSON client", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("surfaces OperationErrors even when HTTP status is 200", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        status: 200,
        headers: new Headers({ TrackingId: "track-1" }),
        text: async () =>
          JSON.stringify({
            TrackingId: "track-1",
            OperationErrors: [
              {
                Code: 100,
                ErrorCode: "NullRequest",
                Message: "The request message is null.",
              },
            ],
          }),
      }),
    );

    await expect(
      microsoftJsonRequest({
        service: "reporting",
        url: "https://reporting.api.bingads.microsoft.com/Reporting/v13/GenerateReport/Submit",
        context: { accessToken: "token", customerId: "11", accountId: "123" },
        body: {},
      }),
    ).rejects.toMatchObject({
      message: "NullRequest: The request message is null.",
    });
  });
});
