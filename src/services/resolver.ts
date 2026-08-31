import {
  AccountNotFoundError,
  AmbiguousAccountError,
  ConnectionNotFoundError,
  CustomerNotFoundError,
} from "@/lib/errors";
import { getAccessTokenForConnection } from "@/services/tokenService";
import { getAppStore } from "@/store/app-store";
import type { MicrosoftAccountRecord, MicrosoftConnectionRecord, MicrosoftCustomerRecord } from "@/store/types";

export interface ResolvedAccountAccess {
  operatorId: string;
  connection: MicrosoftConnectionRecord;
  customer: MicrosoftCustomerRecord | undefined;
  account: MicrosoftAccountRecord;
  accessToken: string;
}

export async function loadCurrentOperatorRecord() {
  const { getOperatorContext } = await import("@/lib/request-context");
  const { operatorId } = getOperatorContext();
  const operator = await getAppStore().getOperator(operatorId);
  if (!operator) {
    throw new ConnectionNotFoundError("The authenticated operator was not found.");
  }
  return operator;
}

export async function resolveOwnedConnection(
  operatorId: string,
  connectionId: string,
): Promise<MicrosoftConnectionRecord> {
  const connection = await getAppStore().getConnection(connectionId);
  if (!connection || connection.operatorId !== operatorId || connection.status === "disconnected") {
    throw new ConnectionNotFoundError();
  }
  return connection;
}

export async function resolveAccountAccess(params: {
  operatorId: string;
  connectionId?: string;
  customerId?: string;
  accountId: string;
}): Promise<ResolvedAccountAccess> {
  const store = getAppStore();
  const candidates = await store.findAccountsByAccountId(params.operatorId, params.accountId);
  const matches: MicrosoftAccountRecord[] = [];
  for (const item of candidates) {
    if (params.connectionId && item.connectionId !== params.connectionId) {
      continue;
    }
    if (params.customerId && item.customerId !== params.customerId) {
      continue;
    }
    const connection = await store.getConnection(item.connectionId);
    if (!connection || connection.operatorId !== params.operatorId || connection.status === "disconnected") {
      continue;
    }
    matches.push(item);
  }

  if (matches.length === 0) {
    throw new AccountNotFoundError();
  }
  if (matches.length > 1) {
    const connectionIds = [...new Set(matches.map((item) => item.connectionId))];
    throw new AmbiguousAccountError(
      `Multiple Microsoft connections have access to account ${params.accountId}. Retry with connectionId set to one of: ${connectionIds.join(", ")}.`,
    );
  }
  const account = matches[0] as MicrosoftAccountRecord;
  const connection = await resolveOwnedConnection(params.operatorId, account.connectionId);
  const customer = await store.getCustomer(
    params.operatorId,
    account.connectionId,
    account.customerId,
  );
  const accessToken = await getAccessTokenForConnection(params.operatorId, connection.connectionId);
  return { operatorId: params.operatorId, connection, customer, account, accessToken };
}

export async function resolveCustomerAccess(params: {
  operatorId: string;
  connectionId?: string;
  customerId: string;
}): Promise<{
  operatorId: string;
  connection: MicrosoftConnectionRecord;
  customer: MicrosoftCustomerRecord;
}> {
  const store = getAppStore();
  const candidates = await store.findCustomersByCustomerId(params.operatorId, params.customerId);
  const matches: MicrosoftCustomerRecord[] = [];
  for (const item of candidates) {
    if (params.connectionId && item.connectionId !== params.connectionId) {
      continue;
    }
    const connection = await store.getConnection(item.connectionId);
    if (!connection || connection.operatorId !== params.operatorId || connection.status === "disconnected") {
      continue;
    }
    matches.push(item);
  }
  if (matches.length === 0) {
    throw new CustomerNotFoundError();
  }
  if (matches.length > 1) {
    const connectionIds = [...new Set(matches.map((item) => item.connectionId))];
    throw new AmbiguousAccountError(
      `Multiple Microsoft connections have access to customer ${params.customerId}. Retry with connectionId set to one of: ${connectionIds.join(", ")}.`,
    );
  }
  const customer = matches[0] as MicrosoftCustomerRecord;
  const connection = await resolveOwnedConnection(params.operatorId, customer.connectionId);
  return { operatorId: params.operatorId, connection, customer };
}

export function applyPaging<T>(items: T[], limit = 100, offset = 0): {
  items: T[];
  total: number;
  limit: number;
  offset: number;
} {
  const safeLimit = Math.min(Math.max(limit, 1), 1000);
  const safeOffset = Math.max(offset, 0);
  return {
    items: items.slice(safeOffset, safeOffset + safeLimit),
    total: items.length,
    limit: safeLimit,
    offset: safeOffset,
  };
}
