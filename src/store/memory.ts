import { randomUUID } from "node:crypto";
import { AppError, ConnectionNotFoundError, ValidationError } from "@/lib/errors";
import type {
  AppStore,
  ConnectionStatus,
  DiscoveredAccount,
  DiscoveredCustomer,
  MicrosoftAccountRecord,
  MicrosoftConnectionRecord,
  MicrosoftCustomerRecord,
  OAuthTransactionRecord,
  OperatorRecord,
  SessionRecord,
  UpsertConnectionInput,
} from "@/store/types";
import { connectionUniqueKey, resourceDocId } from "@/store/types";

function nowIso(): string {
  return new Date().toISOString();
}

export function createMemoryAppStore(): AppStore {
  const operators = new Map<string, OperatorRecord>();
  const connections = new Map<string, MicrosoftConnectionRecord>();
  const uniques = new Map<string, string>();
  const customers = new Map<string, MicrosoftCustomerRecord>();
  const accounts = new Map<string, MicrosoftAccountRecord>();
  const transactions = new Map<string, OAuthTransactionRecord>();
  const sessions = new Map<string, SessionRecord>();

  return {
    async createOperator(primaryEmail = null) {
      const timestamp = nowIso();
      const record: OperatorRecord = {
        operatorId: randomUUID(),
        primaryEmail,
        createdAt: timestamp,
        updatedAt: timestamp,
        lastAccessAt: timestamp,
      };
      operators.set(record.operatorId, record);
      return record;
    },

    async getOperator(operatorId) {
      return operators.get(operatorId);
    },

    async touchLastAccess(operatorId) {
      const existing = operators.get(operatorId);
      if (!existing) {
        return;
      }
      operators.set(operatorId, { ...existing, lastAccessAt: nowIso() });
    },

    async getConnection(connectionId) {
      return connections.get(connectionId);
    },

    async getConnectionForOperatorMicrosoft(operatorId, microsoftSubjectId) {
      const connectionId = uniques.get(connectionUniqueKey(operatorId, microsoftSubjectId));
      return connectionId ? connections.get(connectionId) : undefined;
    },

    async listConnections(operatorId) {
      return [...connections.values()].filter((item) => item.operatorId === operatorId);
    },

    async listConnectionsByMicrosoftSubject(microsoftSubjectId) {
      return [...connections.values()].filter((item) => item.microsoftSubjectId === microsoftSubjectId);
    },

    async upsertConnection(input: UpsertConnectionInput) {
      const unique = connectionUniqueKey(input.operatorId, input.microsoftSubjectId);
      const existingId = uniques.get(unique);
      const existing = existingId ? connections.get(existingId) : undefined;
      if (existing && input.connectionId && input.connectionId !== existing.connectionId) {
        throw new ValidationError("This Microsoft account is already connected for this operator.");
      }

      const timestamp = nowIso();
      const nextToken = input.encryptedRefreshToken || existing?.refreshTokenEncrypted || null;
      if (!nextToken) {
        throw new AppError(
          "Microsoft did not return a refresh token. Reauthorize this account and grant offline access.",
          "revoked",
          400,
        );
      }

      const record: MicrosoftConnectionRecord = {
        connectionId: existing?.connectionId ?? input.connectionId ?? randomUUID(),
        operatorId: input.operatorId,
        microsoftSubjectId: input.microsoftSubjectId,
        email: input.email ?? existing?.email ?? null,
        displayName: input.displayName ?? existing?.displayName ?? null,
        refreshTokenEncrypted: nextToken,
        scopes: input.scopes ?? existing?.scopes ?? [],
        status: "active",
        createdAt: existing?.createdAt ?? timestamp,
        updatedAt: timestamp,
        lastUsedAt: existing?.lastUsedAt ?? null,
        lastSyncedAt: existing?.lastSyncedAt ?? null,
      };
      connections.set(record.connectionId, record);
      uniques.set(unique, record.connectionId);
      return record;
    },

    async markConnectionStatus(connectionId, status: ConnectionStatus, extra) {
      const existing = connections.get(connectionId);
      if (!existing) {
        throw new ConnectionNotFoundError();
      }
      const next: MicrosoftConnectionRecord = {
        ...existing,
        status,
        updatedAt: nowIso(),
        lastUsedAt: extra?.lastUsedAt ?? existing.lastUsedAt,
        lastSyncedAt: extra?.lastSyncedAt ?? existing.lastSyncedAt,
      };
      connections.set(connectionId, next);
      return next;
    },

    async disconnectConnection(operatorId, connectionId) {
      const existing = connections.get(connectionId);
      if (!existing || existing.operatorId !== operatorId) {
        throw new ConnectionNotFoundError();
      }
      const next: MicrosoftConnectionRecord = {
        ...existing,
        refreshTokenEncrypted: null,
        status: "disconnected",
        updatedAt: nowIso(),
      };
      connections.set(connectionId, next);
      return next;
    },

    async replaceConnectionResources(
      operatorId,
      connectionId,
      discoveredCustomers,
      discoveredAccounts,
    ) {
      const timestamp = nowIso();
      for (const [id, customer] of customers) {
        if (customer.connectionId === connectionId) {
          customers.delete(id);
        }
      }
      for (const [id, account] of accounts) {
        if (account.connectionId === connectionId) {
          accounts.delete(id);
        }
      }
      for (const customer of discoveredCustomers as DiscoveredCustomer[]) {
        const id = resourceDocId(connectionId, customer.customerId);
        customers.set(id, {
          id,
          connectionId,
          operatorId,
          ...customer,
          createdAt: timestamp,
          updatedAt: timestamp,
          lastSyncedAt: timestamp,
        });
      }
      for (const account of discoveredAccounts as DiscoveredAccount[]) {
        const id = resourceDocId(connectionId, account.accountId);
        accounts.set(id, {
          id,
          connectionId,
          operatorId,
          ...account,
          createdAt: timestamp,
          updatedAt: timestamp,
          lastSyncedAt: timestamp,
        });
      }
      const connection = connections.get(connectionId);
      if (connection) {
        connections.set(connectionId, {
          ...connection,
          lastSyncedAt: timestamp,
          updatedAt: timestamp,
        });
      }
    },

    async listCustomers(operatorId, connectionId) {
      return [...customers.values()].filter(
        (item) =>
          item.operatorId === operatorId && (!connectionId || item.connectionId === connectionId),
      );
    },

    async findCustomersByCustomerId(operatorId, customerId) {
      return [...customers.values()].filter(
        (item) => item.operatorId === operatorId && item.customerId === customerId,
      );
    },

    async getCustomer(operatorId, connectionId, customerId) {
      return [...customers.values()].find(
        (item) =>
          item.operatorId === operatorId &&
          item.connectionId === connectionId &&
          item.customerId === customerId,
      );
    },

    async listAccounts(operatorId, connectionId) {
      return [...accounts.values()].filter(
        (item) =>
          item.operatorId === operatorId && (!connectionId || item.connectionId === connectionId),
      );
    },

    async findAccountsByAccountId(operatorId, accountId) {
      return [...accounts.values()].filter(
        (item) => item.operatorId === operatorId && item.accountId === accountId,
      );
    },

    async getAccount(operatorId, connectionId, accountId) {
      return [...accounts.values()].find(
        (item) =>
          item.operatorId === operatorId &&
          item.connectionId === connectionId &&
          item.accountId === accountId,
      );
    },

    async createTransaction(input) {
      const record: OAuthTransactionRecord = { ...input, consumedAt: null };
      transactions.set(record.nonce, record);
      return record;
    },

    async getTransaction(nonce) {
      return transactions.get(nonce);
    },

    async consumeTransaction(nonce) {
      const existing = transactions.get(nonce);
      if (!existing) {
        throw new AppError("OAuth state is invalid.", "unauthorized", 401);
      }
      if (existing.consumedAt) {
        throw new AppError("OAuth state has already been used.", "unauthorized", 401);
      }
      if (existing.exp < Date.now()) {
        throw new AppError(
          "OAuth state has expired. Start the Microsoft authorization flow again.",
          "unauthorized",
          401,
        );
      }
      const consumed: OAuthTransactionRecord = { ...existing, consumedAt: nowIso() };
      transactions.set(nonce, consumed);
      return consumed;
    },

    async getSession(sessionId) {
      return sessions.get(sessionId);
    },

    async upsertSession(session) {
      sessions.set(session.sessionId, session);
      return session;
    },
  };
}
