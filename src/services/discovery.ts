import { logger } from "@/lib/logger";
import { getUser, getCustomersInfo, searchAccounts } from "@/microsoft/customer-management/api";
import { getAppStore } from "@/store/app-store";
import type { DiscoveredAccount, DiscoveredCustomer } from "@/store/types";

export async function discoverAndPersistResources(
  operatorId: string,
  connectionId: string,
  accessToken: string,
): Promise<{ customers: DiscoveredCustomer[]; accounts: DiscoveredAccount[] }> {
  const context = { accessToken };
  const { user, roles } = await getUser(context);

  let customers: DiscoveredCustomer[] = [];
  try {
    customers = await getCustomersInfo(context);
  } catch (error) {
    logger.warn("GetCustomersInfo failed; falling back to CustomerRoles", {
      operatorId,
      connectionId,
      errorMessage: error instanceof Error ? error.message : "unknown",
    });
    customers = roles
      .filter((role) => role.customerId)
      .map((role) => ({
        customerId: role.customerId,
        customerName: role.customerId,
        customerNumber: null,
        status: null,
      }));
  }

  let accounts: DiscoveredAccount[] = [];
  if (user.id) {
    try {
      accounts = await searchAccounts(context, user.id);
    } catch (error) {
      logger.warn("SearchAccounts failed; falling back to CustomerRoles account IDs", {
        operatorId,
        connectionId,
        errorMessage: error instanceof Error ? error.message : "unknown",
      });
    }
  }

  if (accounts.length === 0) {
    for (const role of roles) {
      for (const accountId of [...role.accountIds, ...role.linkedAccountIds]) {
        accounts.push({
          customerId: role.customerId,
          accountId,
          accountName: accountId,
          accountNumber: null,
          status: null,
          currencyCode: null,
          timeZone: null,
          accountType: null,
        });
      }
    }
  }

  const customerIds = new Set(customers.map((item) => item.customerId));
  for (const account of accounts) {
    if (account.customerId && !customerIds.has(account.customerId)) {
      customers.push({
        customerId: account.customerId,
        customerName: account.customerId,
        customerNumber: null,
        status: null,
      });
      customerIds.add(account.customerId);
    }
  }

  await getAppStore().replaceConnectionResources(operatorId, connectionId, customers, accounts);
  logger.info("Persisted Microsoft Advertising discovery", {
    operatorId,
    connectionId,
    customerCount: customers.length,
    accountCount: accounts.length,
  });
  return { customers, accounts };
}
