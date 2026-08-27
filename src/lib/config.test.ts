import { afterEach, describe, expect, it } from "vitest";
import { getConfig, resolveFirestoreDatabaseId } from "@/lib/config";
import { setRequiredEnv } from "@/test/env";

describe("firestore database id", () => {
  afterEach(() => {
    setRequiredEnv();
  });

  it("requires the dedicated named database and rejects foreign IDs", () => {
    expect(resolveFirestoreDatabaseId("bing-mcp-v1")).toBe("bing-mcp-v1");
    expect(() => resolveFirestoreDatabaseId("(default)")).toThrow(/named database/);
    expect(() => resolveFirestoreDatabaseId("default")).toThrow(/named database/);
    expect(() => resolveFirestoreDatabaseId("gconnect-mcp-all-v1")).toThrow(/bing-mcp-v1/);
    expect(() => resolveFirestoreDatabaseId("")).toThrow(/FIRESTORE_DATABASE_ID/);
    expect(() => resolveFirestoreDatabaseId("something-else")).toThrow(/bing-mcp-v1/);
  });

  it("loads FIRESTORE_DATABASE_ID from env", () => {
    setRequiredEnv();
    expect(getConfig().firestoreDatabaseId).toBe("bing-mcp-v1");
  });
});
