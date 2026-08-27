import { encryptRefreshToken } from "@/lib/crypto";
import { logger } from "@/lib/logger";
import { microsoftScopes, type MicrosoftIdentity } from "@/microsoft/auth/oauth";
import { discoverAndPersistResources } from "@/services/discovery";
import { getAccessTokenForConnection } from "@/services/tokenService";
import { getAppStore } from "@/store/app-store";
import type { MicrosoftConnectionRecord, OperatorRecord } from "@/store/types";

export async function bindMicrosoftIdentityToOperator(params: {
  operatorId: string;
  identity: MicrosoftIdentity;
  encryptedRefreshToken?: string | null;
}): Promise<{ operator: OperatorRecord; connection: MicrosoftConnectionRecord }> {
  const store = getAppStore();
  const operator = await store.getOperator(params.operatorId);
  if (!operator) {
    throw new Error("The operator no longer exists.");
  }

  const encrypted =
    params.identity.refreshToken
      ? encryptRefreshToken(params.identity.refreshToken)
      : params.encryptedRefreshToken || undefined;

  const connection = await store.upsertConnection({
    operatorId: params.operatorId,
    microsoftSubjectId: params.identity.microsoftSubjectId,
    email: params.identity.email,
    displayName: params.identity.displayName,
    encryptedRefreshToken: encrypted,
    scopes: params.identity.scopes.length > 0 ? params.identity.scopes : microsoftScopes(),
  });

  if (connection.refreshTokenEncrypted) {
    try {
      const accessToken = await getAccessTokenForConnection(params.operatorId, connection.connectionId);
      await discoverAndPersistResources(params.operatorId, connection.connectionId, accessToken);
    } catch (error) {
      logger.warn("Microsoft Advertising discovery failed after bind", {
        operatorId: params.operatorId,
        connectionId: connection.connectionId,
        errorCategory: error instanceof Error ? error.name : "unknown",
        errorMessage: error instanceof Error ? error.message : "unknown",
      });
    }
  }

  return { operator, connection };
}

export async function createOperatorWithConnection(
  identity: MicrosoftIdentity,
  encryptedRefreshToken?: string | null,
) {
  const operator = await getAppStore().createOperator(identity.email);
  return bindMicrosoftIdentityToOperator({
    operatorId: operator.operatorId,
    identity,
    encryptedRefreshToken,
  });
}
