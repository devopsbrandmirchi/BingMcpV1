import { ValidationError } from "@/lib/errors";

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

export function isIsoDate(value: string): boolean {
  const match = value.match(ISO_DATE);
  if (!match) {
    return false;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

export function parseIsoDate(value: string, fieldName: string): string {
  const trimmed = value.trim();
  if (!isIsoDate(trimmed)) {
    throw new ValidationError(`${fieldName} must be a calendar date in YYYY-MM-DD format.`);
  }
  return trimmed;
}

export function assertDateRange(startDate: string, endDate: string): { startDate: string; endDate: string } {
  const start = parseIsoDate(startDate, "startDate");
  const end = parseIsoDate(endDate, "endDate");
  if (start > end) {
    throw new ValidationError("startDate must be on or before endDate.");
  }
  return { startDate: start, endDate: end };
}

export function toMicrosoftReportDate(isoDate: string): { Year: number; Month: number; Day: number } {
  const [year, month, day] = isoDate.split("-").map(Number);
  return { Year: year as number, Month: month as number, Day: day as number };
}

export const REPORT_TIME_PERIODS = [
  "Today",
  "Yesterday",
  "LastSevenDays",
  "LastThirtyDays",
  "ThisWeek",
  "LastWeek",
  "ThisMonth",
  "LastMonth",
] as const;

export type ReportTimePeriod = (typeof REPORT_TIME_PERIODS)[number];

export function isReportTimePeriod(value: string): value is ReportTimePeriod {
  return (REPORT_TIME_PERIODS as readonly string[]).includes(value);
}

export function assertReportWindow(input: {
  period?: string;
  startDate?: string;
  endDate?: string;
}): { period?: ReportTimePeriod; startDate?: string; endDate?: string } {
  if (input.period) {
    if (!isReportTimePeriod(input.period)) {
      throw new ValidationError(`period must be one of: ${REPORT_TIME_PERIODS.join(", ")}.`);
    }
    return { period: input.period };
  }
  if (input.startDate && input.endDate) {
    return assertDateRange(input.startDate, input.endDate);
  }
  throw new ValidationError(
    "Provide period (ThisMonth, ThisWeek, LastSevenDays, ...) or both startDate and endDate as YYYY-MM-DD.",
  );
}
