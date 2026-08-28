import { peekOperatorContext } from "@/lib/request-context";
import { logger } from "@/lib/logger";
import { MicrosoftApiError, MicrosoftRateLimitError, ValidationError, mapMicrosoftError } from "@/lib/errors";
import { getConfig } from "@/lib/config";
import type { MicrosoftRequestContext, MicrosoftService } from "@/microsoft/client/headers";
import { childText, stripXmlNamespaces, xmlEscape } from "@/microsoft/client/xml";

const MAX_RETRIES = 3;
const BASE_DELAY_MS = 500;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isRetryableStatus(status: number): boolean {
  return status === 429 || status === 502 || status === 503 || status === 504;
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

function soapErrorFromXml(xml: string): { code?: number; message: string } | null {
  const stripped = stripXmlNamespaces(xml);
  if (!/<Fault[\s>]|<ApiFaultDetail[\s>]|<AdApiFaultDetail[\s>]/i.test(stripped)) {
    return null;
  }
  const errorCode = childText(stripped, "ErrorCode");
  const message = childText(stripped, "Message") ?? childText(stripped, "faultstring") ?? "Microsoft Advertising SOAP request failed.";
  const codeText = childText(stripped, "Code");
  const code = codeText ? Number(codeText) : undefined;
  return {
    code: Number.isFinite(code) ? code : undefined,
    message: errorCode ? `${errorCode}: ${message}` : message,
  };
}

function trackingIdFromXml(xml: string, response: Response): string | undefined {
  return childText(stripXmlNamespaces(xml), "TrackingId") ?? response.headers.get("TrackingId") ?? undefined;
}

export function buildSoapEnvelope(params: {
  action: string;
  namespace: string;
  context: MicrosoftRequestContext;
  bodyXml: string;
}): string {
  if (!params.context.customerId) {
    throw new ValidationError(
      "A Microsoft Advertising customer ID is required for campaign and reporting calls. Reconnect the Microsoft account to refresh customers and accounts.",
    );
  }

  return [
    `<?xml version="1.0" encoding="utf-8"?>`,
    `<s:Envelope xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">`,
    `<s:Header xmlns="${params.namespace}">`,
    `<Action mustUnderstand="1">${xmlEscape(params.action)}</Action>`,
    `<AuthenticationToken>${xmlEscape(params.context.accessToken)}</AuthenticationToken>`,
    params.context.accountId
      ? `<CustomerAccountId>${xmlEscape(params.context.accountId)}</CustomerAccountId>`
      : "",
    `<CustomerId>${xmlEscape(params.context.customerId)}</CustomerId>`,
    `<DeveloperToken>${xmlEscape(getConfig().microsoftAdsDeveloperToken)}</DeveloperToken>`,
    `</s:Header>`,
    `<s:Body>`,
    params.bodyXml,
    `</s:Body>`,
    `</s:Envelope>`,
  ]
    .filter(Boolean)
    .join("");
}

export async function microsoftSoapRequest(params: {
  service: MicrosoftService;
  url: string;
  action: string;
  namespace: string;
  context: MicrosoftRequestContext;
  bodyXml: string;
}): Promise<string> {
  const requestId = peekOperatorContext()?.requestId;
  const envelope = buildSoapEnvelope({
    action: params.action,
    namespace: params.namespace,
    context: params.context,
    bodyXml: params.bodyXml,
  });
  let lastError: unknown;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    const response = await fetch(params.url, {
      method: "POST",
      headers: {
        "Content-Type": "text/xml; charset=utf-8",
        SOAPAction: `"${params.action}"`,
        ...(requestId ? { "X-Request-Id": requestId } : {}),
      },
      body: envelope,
    });

    const xml = await response.text();
    const trackingId = trackingIdFromXml(xml, response);
    const soapError = soapErrorFromXml(xml);

    logger.info("Microsoft Advertising SOAP call", {
      requestId,
      service: params.service,
      url: params.url,
      action: params.action,
      status: response.status,
      trackingId,
      attempt,
    });

    if (response.ok && !soapError) {
      return xml;
    }

    const message = soapError?.message ?? `Microsoft Advertising SOAP returned HTTP ${response.status}.`;
    logger.warn("Microsoft Advertising SOAP error", {
      requestId,
      service: params.service,
      url: params.url,
      action: params.action,
      status: response.status,
      trackingId,
      errorCode: soapError?.code,
      message,
    });

    if (soapError?.code === 117 || soapError?.code === 207 || response.status === 429) {
      lastError = new MicrosoftRateLimitError(message);
      if (attempt < MAX_RETRIES) {
        await sleep(retryAfterMs(response, attempt));
        continue;
      }
      throw lastError;
    }

    if (!soapError && isRetryableStatus(response.status) && attempt < MAX_RETRIES) {
      lastError = new MicrosoftApiError(message, response.status, trackingId);
      await sleep(retryAfterMs(response, attempt));
      continue;
    }

    throw mapMicrosoftError({
      status: response.status,
      code: soapError?.code ?? response.status,
      message,
      trackingId,
    });
  }

  throw mapMicrosoftError(lastError);
}
