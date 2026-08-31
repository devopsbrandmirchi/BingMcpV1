import { beforeEach, describe, expect, it, vi } from "vitest";
import { encryptRefreshToken } from "@/lib/crypto";
import { runWithOperator } from "@/lib/request-context";
import { issueAccessToken, readAccessToken } from "@/mcp/oauth/tokens";
import { ConnectionNotFoundError } from "@/lib/errors";
import { resolveAccountAccess } from "@/services/resolver";
import { createMemoryAppStore } from "@/store/memory";
import { resetAppStore, setAppStore } from "@/store/app-store";
import { setRequiredEnv } from "@/test/env";

vi.mock("@/services/tokenService", () => ({
  getAccessTokenForConnection: vi.fn(async () => "access-token"),
}));

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
      ).rejects.toMatchObject({
        name: "AmbiguousAccountError",
        message: expect.stringContaining(connectionX.connectionId),
      });
    });
  });

  it("ignores disconnected connections when resolving an account", async () => {
    const { getAppStore } = await import("@/store/app-store");
    const store = getAppStore();
    const operator = await store.createOperator("a@example.com");
    const live = await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-live",
      encryptedRefreshToken: encryptRefreshToken("rl"),
    });
    const stale = await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-stale",
      encryptedRefreshToken: encryptRefreshToken("rs"),
    });
    const account = {
      customerId: "c1",
      accountId: "188405633",
      accountName: "Zoomers RV",
      accountNumber: "G12068D5",
      status: "Active",
      currencyCode: "USD",
      timeZone: "CentralTimeUSCanada",
      accountType: "Advertiser",
    };
    await store.replaceConnectionResources(
      operator.operatorId,
      live.connectionId,
      [{ customerId: "c1", customerName: "C1", customerNumber: null, status: "Active" }],
      [account],
    );
    await store.replaceConnectionResources(
      operator.operatorId,
      stale.connectionId,
      [{ customerId: "c1", customerName: "C1", customerNumber: null, status: "Active" }],
      [account],
    );
    await store.disconnectConnection(operator.operatorId, stale.connectionId);

    await runWithOperator({ requestId: "r2", operatorId: operator.operatorId }, async () => {
      const resolved = await resolveAccountAccess({
        operatorId: operator.operatorId,
        accountId: "188405633",
      });
      expect(resolved.connection.connectionId).toBe(live.connectionId);
    });
  });
});
