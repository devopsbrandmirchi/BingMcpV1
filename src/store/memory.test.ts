import { describe, expect, it } from "vitest";
import { encryptRefreshToken } from "@/lib/crypto";
import { createMemoryAppStore } from "@/store/memory";
import { setRequiredEnv } from "@/test/env";

describe("memory store", () => {
  it("creates operators and connections", async () => {
    setRequiredEnv();
    const store = createMemoryAppStore();
    const operator = await store.createOperator("a@example.com");
    const connection = await store.upsertConnection({
      operatorId: operator.operatorId,
      microsoftSubjectId: "ms-1",
      email: "a@example.com",
      encryptedRefreshToken: encryptRefreshToken("refresh"),
    });
    expect(connection.status).toBe("active");
    const listed = await store.listConnections(operator.operatorId);
    expect(listed).toHaveLength(1);
  });
});
