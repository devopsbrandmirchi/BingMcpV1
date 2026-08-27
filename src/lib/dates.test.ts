import { describe, expect, it } from "vitest";
import { assertDateRange, isIsoDate } from "@/lib/dates";

describe("date validation", () => {
  it("accepts valid YYYY-MM-DD dates and rejects invalid ones", () => {
    expect(isIsoDate("2026-08-01")).toBe(true);
    expect(isIsoDate("2026-13-01")).toBe(false);
    expect(isIsoDate("08/01/2026")).toBe(false);
    expect(isIsoDate("2026-8-1")).toBe(false);
  });

  it("rejects inverted ranges", () => {
    expect(() => assertDateRange("2026-08-20", "2026-08-01")).toThrow(/startDate/);
    expect(assertDateRange("2026-08-01", "2026-08-20")).toEqual({
      startDate: "2026-08-01",
      endDate: "2026-08-20",
    });
  });
});
