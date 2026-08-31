import { z } from "zod";
import { isIsoDate, REPORT_TIME_PERIODS } from "@/lib/dates";

export const connectionIdSchema = z.string().min(1).max(128);
export const optionalConnectionIdSchema = connectionIdSchema.optional();
export const microsoftIdSchema = z.string().min(1).max(64);
export const optionalMicrosoftIdSchema = microsoftIdSchema.optional();

export const isoDateSchema = z.string().refine(isIsoDate, {
  message: "Use a calendar date in YYYY-MM-DD format.",
});

export const reportPeriodSchema = z.enum(REPORT_TIME_PERIODS);

export const pagingSchema = {
  limit: z.number().int().min(1).max(1000).default(100),
  offset: z.number().int().min(0).max(100000).default(0),
};

export const optionalRefreshSchema = z.boolean().optional();

export function jsonToolResult(data: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(data, null, 2),
      },
    ],
  };
}
