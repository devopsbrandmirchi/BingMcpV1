import { ValidationError } from "@/lib/errors";

export function toMicrosoftLong(value: string, fieldName = "id"): number {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) {
    throw new ValidationError(`${fieldName} must be a numeric Microsoft Advertising identifier.`);
  }
  return parsed;
}

export function toMicrosoftLongs(values: string[], fieldName = "ids"): number[] {
  return values.map((value) => toMicrosoftLong(value, fieldName));
}
