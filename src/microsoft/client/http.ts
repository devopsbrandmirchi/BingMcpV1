import { peekOperatorContext } from "@/lib/request-context";
import { logger } from "@/lib/logger";
import { MicrosoftApiError, MicrosoftRateLimitError, ValidationError, mapMicrosoftError } from "@/lib/errors";
import { buildMicrosoftHeaders, type MicrosoftRequestContext, type MicrosoftService } from "@/microsoft/client/headers";

const MAX_RETRIES = 3;
const BASE_DELAY_MS = 500;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryableStatus(status: number): boolean {
  return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

function retryAfterMs(response: Response, attempt: number): number {
  const header = response.headers.get("retry-after");
  if (header) {
    const seconds = Number(header);
    if (Number.isFinite(seconds) && seconds >= 0) {
      return seconds * 1000;
    }
  }
  return BASE_DELAY_MS * 2 ** attempt;
}

function extractTrackingId(body: unknown, response: Response): string | undefined {
  if (body && typeof body === "object" && "TrackingId" in body && typeof body.TrackingId === "string") {
    return body.TrackingId;
  }
  return response.headers.get("TrackingId") ?? undefined;
}

interface MicrosoftErrorItem {
  Code?: number;
  ErrorCode?: string;
  Message?: string;
}

function microsoftErrorItems(body: unknown): MicrosoftErrorItem[] {
  if (!body || typeof body !== "object") {
    return [];
  }
  const record = body as Record<string, unknown>;
  const buckets = [record.Errors, record.OperationErrors, record.BatchErrors];
  const items: MicrosoftErrorItem[] = [];
  for (const bucket of buckets) {
    if (Array.isArray(bucket)) {
      items.push(...(bucket as MicrosoftErrorItem[]));
    }
  }
  return items;
}

function microsoftErrorCode(body: unknown): number | undefined {
  return microsoftErrorItems(body)[0]?.Code;
}

function microsoftErrorMessage(body: unknown, fallback: string): string {
  const first = microsoftErrorItems(body)[0];
  if (first?.Message) {
    return first.ErrorCode ? `${first.ErrorCode}: ${first.Message}` : first.Message;
  }
  return fallback;
}

function hasFatalMicrosoftErrors(body: unknown): boolean {
  return microsoftErrorItems(body).length > 0;
}

export async function microsoftJsonRequest<T>(params: {
  service: MicrosoftService;
  url: string;
  context: MicrosoftRequestContext;
  body?: unknown;
}): Promise<T> {
  if (
    (params.service === "campaign-management" || params.service === "reporting") &&
    !params.context.customerId
  ) {
    throw new ValidationError(
      "A Microsoft Advertising customer ID is required for campaign and reporting calls. Reconnect the Microsoft account to refresh customers and accounts.",
    );
  }

  const requestId = peekOperatorContext()?.requestId;
  let lastError: unknown;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    const response = await fetch(params.url, {
      method: "POST",
      headers: {
        ...buildMicrosoftHeaders(params.service, params.context),
        ...(requestId ? { "X-Request-Id": requestId } : {}),
      },
      body: JSON.stringify(params.body ?? {}),
    });

    const text = await response.text();
    let parsed: unknown;
    try {
      parsed = text ? JSON.parse(text) : {};
    } catch {
      parsed = { message: text };
    }

    const trackingId = extractTrackingId(parsed, response);
    logger.info("Microsoft Advertising API call", {
      requestId,
      service: params.service,
      url: params.url,
      status: response.status,
      trackingId,
      attempt,
    });

    if (response.ok && !hasFatalMicrosoftErrors(parsed)) {
      return parsed as T;
    }

    const errorCode = microsoftErrorCode(parsed);
    const message = microsoftErrorMessage(parsed, `Microsoft Advertising API returned HTTP ${response.status}.`);
    logger.warn("Microsoft Advertising API error", {
      requestId,
      service: params.service,
      url: params.url,
      status: response.status,
      trackingId,
      errorCode,
      message,
    });

    if (errorCode === 117 || errorCode === 207 || response.status === 429) {
      lastError = new MicrosoftRateLimitError(message);
      if (attempt < MAX_RETRIES) {
        await sleep(retryAfterMs(response, attempt));
        continue;
      }
      throw lastError;
    }

    if (isRetryableStatus(response.status) && attempt < MAX_RETRIES) {
      lastError = new MicrosoftApiError(message, response.status, trackingId);
      await sleep(retryAfterMs(response, attempt));
      continue;
    }

    throw mapMicrosoftError({
      status: response.status,
      code: errorCode ?? response.status,
      message,
      trackingId,
    });
  }

  throw mapMicrosoftError(lastError);
}
