import { Firestore } from "@google-cloud/firestore";
import { randomUUID } from "node:crypto";
import { getConfig } from "@/lib/config";
import { AppError, ConnectionNotFoundError } from "@/lib/errors";
import type {
  AppStore,
  ConnectionStatus,
  MicrosoftAccountRecord,
  MicrosoftConnectionRecord,
  MicrosoftCustomerRecord,
  OAuthTransactionRecord,
  OperatorRecord,
  SessionRecord,
  UpsertConnectionInput,
} from "@/store/types";
import { connectionUniqueKey, resourceDocId } from "@/store/types";

const OPERATORS = "bing_mcp_operators";
const CONNECTIONS = "bing_mcp_microsoft_connections";
const UNIQUES = "bing_mcp_connection_uniques";
const CUSTOMERS = "bing_mcp_microsoft_customers";
const ACCOUNTS = "bing_mcp_microsoft_accounts";
const TRANSACTIONS = "bing_mcp_oauth_transactions";
const SESSIONS = "bing_mcp_sessions";

function nowIso(): string {
  return new Date().toISOString();
}

export function createFirestoreClient(): Firestore {
  const config = getConfig();
  return new Firestore({
    projectId: config.firestoreProjectId,
    databaseId: config.firestoreDatabaseId,
    ignoreUndefinedProperties: true,
  });
}

export function createFirestoreAppStore(dbFactory?: () => Firestore): AppStore {
  let client: Firestore | undefined;
  const db = () => {
    if (!client) {
      client = dbFactory ? dbFactory() : createFirestoreClient();
    }
    return client;
  };

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
      await db().collection(OPERATORS).doc(record.operatorId).set(record);
      return record;
    },

    async getOperator(operatorId) {
      const snapshot = await db().collection(OPERATORS).doc(operatorId).get();
      return snapshot.exists ? (snapshot.data() as OperatorRecord) : undefined;
    },

    async touchLastAccess(operatorId) {
      const ref = db().collection(OPERATORS).doc(operatorId);
      await db().runTransaction(async (tx) => {
        const snapshot = await tx.get(ref);
        if (!snapshot.exists) {
          return;
        }
        tx.update(ref, { lastAccessAt: nowIso() });
      });
    },

    async getConnection(connectionId) {
      const snapshot = await db().collection(CONNECTIONS).doc(connectionId).get();
      return snapshot.exists ? (snapshot.data() as MicrosoftConnectionRecord) : undefined;
    },

    async getConnectionForOperatorMicrosoft(operatorId, microsoftSubjectId) {
      const unique = await db()
        .collection(UNIQUES)
        .doc(connectionUniqueKey(operatorId, microsoftSubjectId))
        .get();
      const connectionId = unique.data()?.connectionId;
      if (typeof connectionId !== "string") {
        return undefined;
      }
      return this.getConnection(connectionId);
    },

    async listConnections(operatorId) {
      const snapshot = await db().collection(CONNECTIONS).where("operatorId", "==", operatorId).get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftConnectionRecord);
    },

    async listConnectionsByMicrosoftSubject(microsoftSubjectId) {
      const snapshot = await db()
        .collection(CONNECTIONS)
        .where("microsoftSubjectId", "==", microsoftSubjectId)
        .get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftConnectionRecord);
    },

    async upsertConnection(input: UpsertConnectionInput) {
      const uniqueKey = connectionUniqueKey(input.operatorId, input.microsoftSubjectId);
      const uniqueRef = db().collection(UNIQUES).doc(uniqueKey);

      return db().runTransaction(async (tx) => {
        const uniqueSnap = await tx.get(uniqueRef);
        const existingId =
          typeof uniqueSnap.data()?.connectionId === "string"
            ? String(uniqueSnap.data()?.connectionId)
            : input.connectionId;
        const connectionRef = existingId
          ? db().collection(CONNECTIONS).doc(existingId)
          : db().collection(CONNECTIONS).doc();
        const existingSnap = existingId ? await tx.get(connectionRef) : undefined;
        const existing = existingSnap?.exists
          ? (existingSnap.data() as MicrosoftConnectionRecord)
          : undefined;

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
          connectionId: existing?.connectionId ?? connectionRef.id,
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
        tx.set(connectionRef, record);
        tx.set(uniqueRef, {
          operatorId: input.operatorId,
          microsoftSubjectId: input.microsoftSubjectId,
          connectionId: record.connectionId,
        });
        return record;
      });
    },

    async markConnectionStatus(connectionId, status: ConnectionStatus, extra) {
      const ref = db().collection(CONNECTIONS).doc(connectionId);
      return db().runTransaction(async (tx) => {
        const snapshot = await tx.get(ref);
        if (!snapshot.exists) {
          throw new ConnectionNotFoundError();
        }
        const existing = snapshot.data() as MicrosoftConnectionRecord;
        const next: MicrosoftConnectionRecord = {
          ...existing,
          status,
          updatedAt: nowIso(),
          lastUsedAt: extra?.lastUsedAt ?? existing.lastUsedAt,
          lastSyncedAt: extra?.lastSyncedAt ?? existing.lastSyncedAt,
        };
        tx.set(ref, next);
        return next;
      });
    },

    async disconnectConnection(operatorId, connectionId) {
      const ref = db().collection(CONNECTIONS).doc(connectionId);
      return db().runTransaction(async (tx) => {
        const snapshot = await tx.get(ref);
        const existing = snapshot.exists
          ? (snapshot.data() as MicrosoftConnectionRecord)
          : undefined;
        if (!existing || existing.operatorId !== operatorId) {
          throw new ConnectionNotFoundError();
        }
        const next: MicrosoftConnectionRecord = {
          ...existing,
          refreshTokenEncrypted: null,
          status: "disconnected",
          updatedAt: nowIso(),
        };
        tx.set(ref, next);
        return next;
      });
    },

    async replaceConnectionResources(operatorId, connectionId, discoveredCustomers, discoveredAccounts) {
      const timestamp = nowIso();
      const customerSnap = await db()
        .collection(CUSTOMERS)
        .where("connectionId", "==", connectionId)
        .get();
      const accountSnap = await db()
        .collection(ACCOUNTS)
        .where("connectionId", "==", connectionId)
        .get();

      const batch = db().batch();
      for (const doc of customerSnap.docs) {
        batch.delete(doc.ref);
      }
      for (const doc of accountSnap.docs) {
        batch.delete(doc.ref);
      }
      for (const customer of discoveredCustomers) {
        const id = resourceDocId(connectionId, customer.customerId);
        batch.set(db().collection(CUSTOMERS).doc(id), {
          id,
          connectionId,
          operatorId,
          ...customer,
          createdAt: timestamp,
          updatedAt: timestamp,
          lastSyncedAt: timestamp,
        } satisfies MicrosoftCustomerRecord);
      }
      for (const account of discoveredAccounts) {
        const id = resourceDocId(connectionId, account.accountId);
        batch.set(db().collection(ACCOUNTS).doc(id), {
          id,
          connectionId,
          operatorId,
          ...account,
          createdAt: timestamp,
          updatedAt: timestamp,
          lastSyncedAt: timestamp,
        } satisfies MicrosoftAccountRecord);
      }
      const connectionRef = db().collection(CONNECTIONS).doc(connectionId);
      batch.set(
        connectionRef,
        { lastSyncedAt: timestamp, updatedAt: timestamp },
        { merge: true },
      );
      await batch.commit();
    },

    async listCustomers(operatorId, connectionId) {
      let query = db().collection(CUSTOMERS).where("operatorId", "==", operatorId);
      if (connectionId) {
        query = query.where("connectionId", "==", connectionId);
      }
      const snapshot = await query.get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftCustomerRecord);
    },

    async findCustomersByCustomerId(operatorId, customerId) {
      const snapshot = await db()
        .collection(CUSTOMERS)
        .where("operatorId", "==", operatorId)
        .where("customerId", "==", customerId)
        .get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftCustomerRecord);
    },

    async getCustomer(operatorId, connectionId, customerId) {
      const snapshot = await db()
        .collection(CUSTOMERS)
        .doc(resourceDocId(connectionId, customerId))
        .get();
      const record = snapshot.exists ? (snapshot.data() as MicrosoftCustomerRecord) : undefined;
      if (!record || record.operatorId !== operatorId) {
        return undefined;
      }
      return record;
    },

    async listAccounts(operatorId, connectionId) {
      let query = db().collection(ACCOUNTS).where("operatorId", "==", operatorId);
      if (connectionId) {
        query = query.where("connectionId", "==", connectionId);
      }
      const snapshot = await query.get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftAccountRecord);
    },

    async findAccountsByAccountId(operatorId, accountId) {
      const snapshot = await db()
        .collection(ACCOUNTS)
        .where("operatorId", "==", operatorId)
        .where("accountId", "==", accountId)
        .get();
      return snapshot.docs.map((doc) => doc.data() as MicrosoftAccountRecord);
    },

    async getAccount(operatorId, connectionId, accountId) {
      const snapshot = await db()
        .collection(ACCOUNTS)
        .doc(resourceDocId(connectionId, accountId))
        .get();
      const record = snapshot.exists ? (snapshot.data() as MicrosoftAccountRecord) : undefined;
      if (!record || record.operatorId !== operatorId) {
        return undefined;
      }
      return record;
    },

    async createTransaction(input) {
      const record: OAuthTransactionRecord = { ...input, consumedAt: null };
      await db().collection(TRANSACTIONS).doc(record.nonce).set(record);
      return record;
    },

    async getTransaction(nonce) {
      const snapshot = await db().collection(TRANSACTIONS).doc(nonce).get();
      return snapshot.exists ? (snapshot.data() as OAuthTransactionRecord) : undefined;
    },

    async consumeTransaction(nonce) {
      const ref = db().collection(TRANSACTIONS).doc(nonce);
      return db().runTransaction(async (tx) => {
        const snapshot = await tx.get(ref);
        if (!snapshot.exists) {
          throw new AppError("OAuth state is invalid.", "unauthorized", 401);
        }
        const existing = snapshot.data() as OAuthTransactionRecord;
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
        tx.set(ref, consumed);
        return consumed;
      });
    },

    async getSession(sessionId) {
      const snapshot = await db().collection(SESSIONS).doc(sessionId).get();
      return snapshot.exists ? (snapshot.data() as SessionRecord) : undefined;
    },

    async upsertSession(session) {
      await db().collection(SESSIONS).doc(session.sessionId).set(session);
      return session;
    },
  };
}
