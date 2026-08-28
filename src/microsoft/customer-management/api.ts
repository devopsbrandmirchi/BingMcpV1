import { microsoftJsonRequest } from "@/microsoft/client/http";
import { CUSTOMER_MANAGEMENT_BASE } from "@/microsoft/client/version";
import type { MicrosoftRequestContext } from "@/microsoft/client/headers";
import type { MicrosoftCustomerRole, MicrosoftUser } from "@/microsoft/models/types";
import type { DiscoveredAccount, DiscoveredCustomer } from "@/store/types";

function asString(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null;
  }
  return String(value);
}

function asStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((item) => String(item)).filter(Boolean);
}

export async function getUser(context: MicrosoftRequestContext): Promise<{
  user: MicrosoftUser;
  roles: MicrosoftCustomerRole[];
}> {
  const response = await microsoftJsonRequest<{
    User?: Record<string, unknown>;
    CustomerRoles?: Array<Record<string, unknown>>;
  }>({
    service: "customer-management",
    url: `${CUSTOMER_MANAGEMENT_BASE}/User/Query`,
    context,
    body: { UserId: null },
  });

  const rawUser = response.User ?? {};
  const contact = (rawUser.ContactInfo ?? {}) as Record<string, unknown>;
  const name = (rawUser.Name ?? {}) as Record<string, unknown>;

  return {
    user: {
      id: asString(rawUser.Id),
      customerId: asString(rawUser.CustomerId),
      userName: asString(rawUser.UserName),
      email: asString(contact.Email),
      firstName: asString(name.FirstName),
      lastName: asString(name.LastName),
      status: asString(rawUser.UserLifeCycleStatus),
    },
    roles: (response.CustomerRoles ?? []).map((role) => ({
      customerId: String(role.CustomerId ?? ""),
      accountIds: asStringArray(role.AccountIds),
      linkedAccountIds: asStringArray(role.LinkedAccountIds),
      roleId: typeof role.RoleId === "number" ? role.RoleId : null,
    })),
  };
}

export async function getCustomersInfo(context: MicrosoftRequestContext): Promise<DiscoveredCustomer[]> {
  const response = await microsoftJsonRequest<{
    CustomersInfo?: Array<Record<string, unknown>>;
  }>({
    service: "customer-management",
    url: `${CUSTOMER_MANAGEMENT_BASE}/CustomersInfo/Query`,
    context,
    body: {},
  });

  return (response.CustomersInfo ?? []).map((item) => ({
    customerId: String(item.Id ?? item.CustomerId ?? ""),
    customerName: String(item.Name ?? ""),
    customerNumber: asString(item.Number),
    status: asString(item.CustomerLifeCycleStatus ?? item.Status),
  })).filter((item) => item.customerId);
}

export async function searchAccounts(
  context: MicrosoftRequestContext,
  userId: string,
): Promise<DiscoveredAccount[]> {
  const response = await microsoftJsonRequest<{
    Accounts?: Array<Record<string, unknown>>;
    AdvertiserAccount?: Array<Record<string, unknown>>;
    AdvertiserAccounts?: Array<Record<string, unknown>>;
  }>({
    service: "customer-management",
    url: `${CUSTOMER_MANAGEMENT_BASE}/Accounts/Search`,
    context,
    body: {
      Predicates: [
        {
          Field: "UserId",
          Operator: "Equals",
          Value: userId,
        },
      ],
      PageInfo: {
        Index: 0,
        Size: 1000,
      },
    },
  });

  const raw = response.Accounts ?? response.AdvertiserAccounts ?? response.AdvertiserAccount ?? [];
  return raw
    .map((item) => ({
      customerId: String(item.ParentCustomerId ?? item.CustomerId ?? ""),
      accountId: String(item.Id ?? ""),
      accountName: String(item.Name ?? ""),
      accountNumber: asString(item.Number),
      status: asString(item.AccountLifeCycleStatus),
      currencyCode: asString(item.CurrencyCode),
      timeZone: asString(item.TimeZone),
      accountType: asString(item.AccountType),
    }))
    .filter((item) => item.accountId);
}

export async function getAccountsInfo(
  context: MicrosoftRequestContext,
  customerId?: string,
): Promise<Array<{ accountId: string; accountName: string; accountNumber: string | null }>> {
  const response = await microsoftJsonRequest<{
    AccountsInfo?: Array<Record<string, unknown>>;
  }>({
    service: "customer-management",
    url: `${CUSTOMER_MANAGEMENT_BASE}/AccountsInfo/Query`,
    context,
    body: {
      CustomerId: customerId ?? null,
      OnlyParentAccounts: false,
    },
  });
  return (response.AccountsInfo ?? []).map((item) => ({
    accountId: String(item.Id ?? ""),
    accountName: String(item.Name ?? ""),
    accountNumber: asString(item.Number),
  })).filter((item) => item.accountId);
}
