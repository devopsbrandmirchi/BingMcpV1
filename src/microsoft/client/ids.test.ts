import { describe, expect, it } from "vitest";
import { toMicrosoftLong, toMicrosoftLongs } from "@/microsoft/client/ids";

describe("Microsoft identifier conversion", () => {
  it("converts numeric string IDs to longs", () => {
    expect(toMicrosoftLong("150419461", "accountId")).toBe(150419461);
    expect(toMicrosoftLongs(["55", "77"], "campaignId")).toEqual([55, 77]);
  });

  it("rejects non-numeric IDs", () => {
    expect(() => toMicrosoftLong("abc", "accountId")).toThrow(/numeric Microsoft Advertising identifier/);
  });
});
