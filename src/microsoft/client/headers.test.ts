import { beforeEach, describe, expect, it } from "vitest";
import { buildMicrosoftHeaders } from "@/microsoft/client/headers";
import { setRequiredEnv } from "@/test/env";

describe("Microsoft request headers", () => {
  beforeEach(() => {
    setRequiredEnv();
  });

  it("omits CustomerId for customer-management", () => {
    const headers = buildMicrosoftHeaders("customer-management", {
      accessToken: "token",
      customerId: "1",
      accountId: "2",
    });
    expect(headers.Authorization).toBe("Bearer token");
    expect(headers.DeveloperToken).toBe("test-ads-developer-token");
    expect(headers.CustomerId).toBeUndefined();
    expect(headers.CustomerAccountId).toBeUndefined();
  });

  it("includes CustomerId and CustomerAccountId for campaign-management", () => {
    const headers = buildMicrosoftHeaders("campaign-management", {
      accessToken: "token",
      customerId: "1",
      accountId: "2",
    });
    expect(headers.CustomerId).toBe("1");
    expect(headers.CustomerAccountId).toBe("2");
  });
});
