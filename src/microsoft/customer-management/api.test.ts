import { beforeEach, describe, expect, it, vi } from "vitest";
import { getUser, searchAccounts } from "@/microsoft/customer-management/api";
import { setRequiredEnv } from "@/test/env";

describe("Customer Management adapters", () => {
  beforeEach(() => {
    setRequiredEnv();
    vi.restoreAllMocks();
  });

  it("maps GetUser and SearchAccounts responses", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn()
        .mockResolvedValueOnce({
          ok: true,
          headers: new Headers(),
          text: async () =>
            JSON.stringify({
              User: {
                Id: "99",
                CustomerId: "11",
                UserName: "user@example.com",
                ContactInfo: { Email: "user@example.com" },
                Name: { FirstName: "Ada", LastName: "Lovelace" },
                UserLifeCycleStatus: "Active",
              },
              CustomerRoles: [{ CustomerId: "11", AccountIds: [123], LinkedAccountIds: [], RoleId: 203 }],
            }),
        })
        .mockResolvedValueOnce({
          ok: true,
          headers: new Headers(),
          text: async () =>
            JSON.stringify({
              Accounts: [
                {
                  Id: 123,
                  Name: "Contoso",
                  Number: "F123",
                  ParentCustomerId: 11,
                  AccountLifeCycleStatus: "Active",
                  CurrencyCode: "USD",
                  TimeZone: "PacificTimeUSCanadaTijuana",
                  AccountType: "Advertiser",
                },
              ],
            }),
        }),
    );

    const user = await getUser({ accessToken: "token" });
    expect(user.user.id).toBe("99");
    expect(user.roles[0]?.customerId).toBe("11");

    const accounts = await searchAccounts({ accessToken: "token" }, "99");
    expect(accounts[0]?.accountId).toBe("123");
    expect(accounts[0]?.accountName).toBe("Contoso");
    expect(accounts[0]?.currencyCode).toBe("USD");
  });
});
