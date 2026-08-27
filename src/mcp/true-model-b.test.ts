import { beforeEach, describe, expect, it } from "vitest";
import { encryptRefreshToken } from "@/lib/crypto";
import { runWithOperator } from "@/lib/request-context";
import { issueAccessToken, readAccessToken } from "@/mcp/oauth/tokens";
import {
  AmbiguousAccountError,
  ConnectionNotFoundError,
} from "@/lib/errors";
import { resolveAccountAccess } from "@/services/resolver";
import { createMemoryAppStore } from "@/store/memory";
import { resetAppStore, setAppStore } from "@/store/app-store";
import { setRequiredEnv } from "@/test/env";

describe("True Model B ownership", () => {
  beforeEach(() => {
    setRequiredEnv();
    resetAppStore();
    setAppStore(createMemoryAppStore());
  });

  it("keeps MCP JWT sub as operatorId", () => {
    const token = issueAccessToken({ clientId: "claude", sub: "operator-a" });
    expect(readAccessToken(token).sub).toBe("operator-a");
  });

  it("allows the same Microsoft identity on two operators independently", async () => {
    const { getAppStore } = await import("@/store/app-store");
    const store = getAppStore();
    const operatorA = await store.createOperator("a@example.com");
    const operatorB = await store.createOperator("b@example.com");
    const encrypted = encryptRefreshToken("refresh-a");
    const connectionA = await store.upsertConnection({
      operatorId: operatorA.operatorId,
      microsoftSubjectId: "ms-same",
      email: "shared@example.com",
      encryptedRefreshToken: encrypted,
    });
    const connectionB = await store.upsertConnection({
      operatorId: operatorB.operatorId,
      microsoftSubjectId: "ms-same",
      email: "shared@example.com",
      encryptedRefreshToken: encryptRefreshToken("refresh-b"),
    });
    expect(connectionA.connectionId).not.toBe(connectionB.connectionId);
    expect(connectionA.operatorId).toBe(operatorA.operatorId);
    expect(connectionB.operatorId).toBe(operatorB.operatorId);
  });

  it("supports multiple Microsoft connections for one operator", async () => {
    const { getAppStore } = await import("@/store/app-store");
    const store = getAppStore();
    const operator = await store.createOperator("a@example.com");
    await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-x",
      email: "x@example.com",
      encryptedRefreshToken: encryptRefreshToken("rx"),
    });
    await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-y",
      email: "y@example.com",
      encryptedRefreshToken: encryptRefreshToken("ry"),
    });
    const connections = await store.listConnections(operator.operatorId);
    expect(connections).toHaveLength(2);
  });

  it("prevents operator A from accessing operator B's connection", async () => {
    const { getAppStore } = await import("@/store/app-store");
    const store = getAppStore();
    const operatorA = await store.createOperator("a@example.com");
    const operatorB = await store.createOperator("b@example.com");
    const connectionB = await store.upsertConnection({
      operatorId: operatorB.operatorId,
      microsoftSubjectId: "ms-b",
      encryptedRefreshToken: encryptRefreshToken("rb"),
    });
    await expect(
      store.disconnectConnection(operatorA.operatorId, connectionB.connectionId),
    ).rejects.toBeInstanceOf(ConnectionNotFoundError);
  });

  it("rejects ambiguous account resolution", async () => {
    const { getAppStore } = await import("@/store/app-store");
    const store = getAppStore();
    const operator = await store.createOperator("a@example.com");
    const connectionX = await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-x",
      encryptedRefreshToken: encryptRefreshToken("rx"),
    });
    const connectionY = await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-y",
      encryptedRefreshToken: encryptRefreshToken("ry"),
    });
    await store.replaceConnectionResources(
      operator.operatorId,
      connectionX.connectionId,
      [{ customerId: "c1", customerName: "C1", customerNumber: null, status: "Active" }],
      [
        {
          customerId: "c1",
          accountId: "123",
          accountName: "Account 123",
          accountNumber: null,
          status: "Active",
          currencyCode: "USD",
          timeZone: "PacificTimeUSCanadaTijuana",
          accountType: "Advertiser",
        },
      ],
    );
    await store.replaceConnectionResources(
      operator.operatorId,
      connectionY.connectionId,
      [{ customerId: "c1", customerName: "C1", customerNumber: null, status: "Active" }],
      [
        {
          customerId: "c1",
          accountId: "123",
          accountName: "Account 123",
          accountNumber: null,
          status: "Active",
          currencyCode: "USD",
          timeZone: "PacificTimeUSCanadaTijuana",
          accountType: "Advertiser",
        },
      ],
    );

    await runWithOperator({ requestId: "r1", operatorId: operator.operatorId }, async () => {
      await expect(
        resolveAccountAccess({
          operatorId: operator.operatorId,
          accountId: "123",
        }),
      ).rejects.toBeInstanceOf(AmbiguousAccountError);
    });
  });
});
