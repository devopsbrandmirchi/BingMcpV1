import type { PendingMcpAuthorize } from "@/mcp/oauth/pending";

export const LEGACY_SESSION_ID = "legacy";

export type ConnectionStatus = "active" | "reauthorization_required" | "disconnected";

export type OAuthOperation =
  | "bootstrap_operator"
  | "add_connection"
  | "reconnect"
  | "resume_choice";

export interface OperatorRecord {
  operatorId: string;
  primaryEmail: string | null;
  createdAt: string;
  updatedAt: string;
  lastAccessAt: string;
}

export interface MicrosoftConnectionRecord {
  connectionId: string;
  operatorId: string;
  microsoftSubjectId: string;
  email: string | null;
  displayName: string | null;
  refreshTokenEncrypted: string | null;
  scopes: string[];
  status: ConnectionStatus;
  createdAt: string;
  updatedAt: string;
  lastUsedAt: string | null;
  lastSyncedAt: string | null;
}

export interface MicrosoftCustomerRecord {
  id: string;
  operatorId: string;
  connectionId: string;
  customerId: string;
  customerName: string;
  customerNumber: string | null;
  status: string | null;
  createdAt: string;
  updatedAt: string;
  lastSyncedAt: string;
}

export interface MicrosoftAccountRecord {
  id: string;
  operatorId: string;
  connectionId: string;
  customerId: string;
  accountId: string;
  accountName: string;
  accountNumber: string | null;
  status: string | null;
  currencyCode: string | null;
  timeZone: string | null;
  accountType: string | null;
  createdAt: string;
  updatedAt: string;
  lastSyncedAt: string;
}

export interface ResumeCandidate {
  operatorId: string;
  primaryEmail: string | null;
  microsoftEmail: string | null;
  connectionCount: number;
  createdAt: string;
}

export interface OAuthTransactionRecord {
  nonce: string;
  operatorId: string | null;
  operation: OAuthOperation;
  connectionId: string | null;
  mcpPending: PendingMcpAuthorize | null;
  codeVerifier: string | null;
  microsoftState: string | null;
  exp: number;
  consumedAt: string | null;
  resumeCandidates?: ResumeCandidate[];
  pendingMicrosoftSubjectId?: string;
  pendingMicrosoftEmail?: string | null;
  pendingMicrosoftDisplayName?: string | null;
  pendingRefreshTokenEncrypted?: string | null;
  pendingMicrosoftScopes?: string[];
}

export interface SessionRecord {
  sessionId: string;
  operatorId: string;
  updatedAt: string;
}

export interface UpsertConnectionInput {
  operatorId: string;
  microsoftSubjectId: string;
  email?: string | null;
  displayName?: string | null;
  encryptedRefreshToken?: string;
  scopes?: string[];
  connectionId?: string;
}

export interface DiscoveredCustomer {
  customerId: string;
  customerName: string;
  customerNumber: string | null;
  status: string | null;
}

export interface DiscoveredAccount {
  customerId: string;
  accountId: string;
  accountName: string;
  accountNumber: string | null;
  status: string | null;
  currencyCode: string | null;
  timeZone: string | null;
  accountType: string | null;
}

export interface AppStore {
  createOperator(primaryEmail?: string | null): Promise<OperatorRecord>;
  getOperator(operatorId: string): Promise<OperatorRecord | undefined>;
  touchLastAccess(operatorId: string): Promise<void>;

  getConnection(connectionId: string): Promise<MicrosoftConnectionRecord | undefined>;
  getConnectionForOperatorMicrosoft(
    operatorId: string,
    microsoftSubjectId: string,
  ): Promise<MicrosoftConnectionRecord | undefined>;
  listConnections(operatorId: string): Promise<MicrosoftConnectionRecord[]>;
  listConnectionsByMicrosoftSubject(
    microsoftSubjectId: string,
  ): Promise<MicrosoftConnectionRecord[]>;
  upsertConnection(input: UpsertConnectionInput): Promise<MicrosoftConnectionRecord>;
  markConnectionStatus(
    connectionId: string,
    status: ConnectionStatus,
    extra?: { lastUsedAt?: string; lastSyncedAt?: string },
  ): Promise<MicrosoftConnectionRecord>;
  disconnectConnection(operatorId: string, connectionId: string): Promise<MicrosoftConnectionRecord>;

  replaceConnectionResources(
    operatorId: string,
    connectionId: string,
    customers: DiscoveredCustomer[],
    accounts: DiscoveredAccount[],
  ): Promise<void>;
  listCustomers(operatorId: string, connectionId?: string): Promise<MicrosoftCustomerRecord[]>;
  findCustomersByCustomerId(
    operatorId: string,
    customerId: string,
  ): Promise<MicrosoftCustomerRecord[]>;
  getCustomer(
    operatorId: string,
    connectionId: string,
    customerId: string,
  ): Promise<MicrosoftCustomerRecord | undefined>;
  listAccounts(operatorId: string, connectionId?: string): Promise<MicrosoftAccountRecord[]>;
  findAccountsByAccountId(operatorId: string, accountId: string): Promise<MicrosoftAccountRecord[]>;
  getAccount(
    operatorId: string,
    connectionId: string,
    accountId: string,
  ): Promise<MicrosoftAccountRecord | undefined>;

  createTransaction(
    input: Omit<OAuthTransactionRecord, "consumedAt">,
  ): Promise<OAuthTransactionRecord>;
  getTransaction(nonce: string): Promise<OAuthTransactionRecord | undefined>;
  consumeTransaction(nonce: string): Promise<OAuthTransactionRecord>;

  getSession(sessionId: string): Promise<SessionRecord | undefined>;
  upsertSession(session: SessionRecord): Promise<SessionRecord>;
}

export function normalizeSessionId(sessionId?: string | null): string {
  const trimmed = sessionId?.trim();
  return trimmed || LEGACY_SESSION_ID;
}

export function resourceDocId(connectionId: string, resourceId: string): string {
  return `${connectionId}_${resourceId.replaceAll("/", "_")}`;
}

export function connectionUniqueKey(operatorId: string, microsoftSubjectId: string): string {
  return `${operatorId}_${microsoftSubjectId}`;
}
