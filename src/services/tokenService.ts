import { decryptRefreshToken, encryptRefreshToken } from "@/lib/crypto";
import {
  ConnectionNotFoundError,
  ConnectionReauthorizationRequiredError,
  mapMicrosoftError,
} from "@/lib/errors";
import { logger } from "@/lib/logger";
import { refreshMicrosoftAccessToken } from "@/microsoft/auth/oauth";
import { getAppStore } from "@/store/app-store";
import type { MicrosoftConnectionRecord } from "@/store/types";

interface CachedAccessToken {
  accessToken: string;
  exp: number;
}

const accessTokenCache = new Map<string, CachedAccessToken>();
const refreshLocks = new Map<string, Promise<string>>();

export async function assertConnectionOwned(
  operatorId: string,
  connectionId: string,
): Promise<MicrosoftConnectionRecord> {
  const connection = await getAppStore().getConnection(connectionId);
  if (!connection || connection.operatorId !== operatorId) {
    throw new ConnectionNotFoundError();
  }
  return connection;
}

async function persistRotatedRefreshToken(
  operatorId: string,
  connection: MicrosoftConnectionRecord,
  refreshToken: string,
  scopes: string[],
): Promise<void> {
  await getAppStore().upsertConnection({
    operatorId,
    microsoftSubjectId: connection.microsoftSubjectId,
    email: connection.email,
    displayName: connection.displayName,
    encryptedRefreshToken: encryptRefreshToken(refreshToken),
    scopes,
    connectionId: connection.connectionId,
  });
}

export async function getAccessTokenForConnection(
  operatorId: string,
  connectionId: string,
): Promise<string> {
  const existingLock = refreshLocks.get(connectionId);
  if (existingLock) {
    return existingLock;
  }

  const work = (async () => {
    const cached = accessTokenCache.get(connectionId);
    if (cached && cached.exp > Date.now() + 60_000) {
      return cached.accessToken;
    }

    const connection = await assertConnectionOwned(operatorId, connectionId);
    if (connection.status === "disconnected") {
      throw new ConnectionNotFoundError("This Microsoft connection has been disconnected.");
    }
    if (connection.status === "reauthorization_required") {
      throw new ConnectionReauthorizationRequiredError(
        `The Microsoft connection for ${connection.email ?? "this account"} requires reauthorization.`,
      );
    }
    if (!connection.refreshTokenEncrypted) {
      throw new ConnectionReauthorizationRequiredError(
        "No Microsoft account is connected to this operator. Connect a Microsoft account first.",
      );
    }

    let refreshToken: string;
    try {
      refreshToken = decryptRefreshToken(connection.refreshTokenEncrypted);
    } catch {
      throw new ConnectionReauthorizationRequiredError(
        "Stored Microsoft credentials could not be decrypted. Reconnect this Microsoft account.",
      );
    }

    try {
      const refreshed = await refreshMicrosoftAccessToken(refreshToken);
      if (refreshed.refreshToken !== refreshToken) {
        await persistRotatedRefreshToken(operatorId, connection, refreshed.refreshToken, refreshed.scopes);
        logger.info("Rotated Microsoft refresh token", { operatorId, connectionId });
      }
      accessTokenCache.set(connectionId, {
        accessToken: refreshed.accessToken,
        exp: refreshed.expiresAt,
      });
      await getAppStore().markConnectionStatus(connectionId, "active", {
        lastUsedAt: new Date().toISOString(),
      });
      return refreshed.accessToken;
    } catch (error) {
      const mapped = mapMicrosoftError(error);
      if (mapped.code === "reauthorization_required" || mapped.code === "revoked") {
        await getAppStore().markConnectionStatus(connectionId, "reauthorization_required");
        throw new ConnectionReauthorizationRequiredError(
          `The Microsoft connection for ${connection.email ?? "this account"} requires reauthorization.`,
        );
      }
      throw mapped;
    }
  })();

  refreshLocks.set(connectionId, work);
  try {
    return await work;
  } finally {
    refreshLocks.delete(connectionId);
  }
}

export function clearAccessTokenCache(connectionId?: string): void {
  if (connectionId) {
    accessTokenCache.delete(connectionId);
    return;
  }
  accessTokenCache.clear();
}

export function publicConnectionView(connection: MicrosoftConnectionRecord) {
  return {
    connectionId: connection.connectionId,
    email: connection.email,
    displayName: connection.displayName,
    status: connection.status,
    lastUsedAt: connection.lastUsedAt,
    lastSyncedAt: connection.lastSyncedAt,
  };
}
