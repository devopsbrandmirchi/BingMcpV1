export type AppErrorCode =
  | "authentication_error"
  | "authorization_error"
  | "unauthorized"
  | "validation"
  | "operator_not_found"
  | "session_invalid"
  | "connection_not_found"
  | "connection_forbidden"
  | "no_microsoft_connections"
  | "reauthorization_required"
  | "customer_not_found"
  | "account_not_found"
  | "account_ambiguous"
  | "customer_ambiguous"
  | "campaign_not_found"
  | "ad_group_not_found"
  | "ad_not_found"
  | "keyword_not_found"
  | "microsoft_api"
  | "microsoft_oauth_error"
  | "rate_limited"
  | "report_error"
  | "report_timeout"
  | "report_empty"
  | "configuration_error"
  | "internal_error"
  | "revoked";

export class AppError extends Error {
  readonly code: AppErrorCode;
  readonly status: number;
  readonly microsoftTrackingId?: string;

  constructor(
    message: string,
    code: AppErrorCode,
    status = 400,
    extras?: { microsoftTrackingId?: string },
  ) {
    super(message);
    this.name = "AppError";
    this.code = code;
    this.status = toHttpStatus(status);
    this.microsoftTrackingId = extras?.microsoftTrackingId;
  }
}

export class AuthenticationError extends AppError {
  constructor(message = "Authentication is required.") {
    super(message, "authentication_error", 401);
    this.name = "AuthenticationError";
  }
}

export class AuthorizationError extends AppError {
  constructor(message = "You are not allowed to access this resource.") {
    super(message, "authorization_error", 403);
    this.name = "AuthorizationError";
  }
}

export class ConnectionNotFoundError extends AppError {
  constructor(message = "The requested Microsoft connection was not found.") {
    super(message, "connection_not_found", 404);
    this.name = "ConnectionNotFoundError";
  }
}

export class ConnectionReauthorizationRequiredError extends AppError {
  constructor(message = "This Microsoft connection needs to be reauthorized.") {
    super(message, "reauthorization_required", 401);
    this.name = "ConnectionReauthorizationRequiredError";
  }
}

export class CustomerNotFoundError extends AppError {
  constructor(message = "The requested Microsoft Advertising customer was not found.") {
    super(message, "customer_not_found", 404);
    this.name = "CustomerNotFoundError";
  }
}

export class AccountNotFoundError extends AppError {
  constructor(message = "The requested Microsoft Advertising account was not found.") {
    super(message, "account_not_found", 404);
    this.name = "AccountNotFoundError";
  }
}

export class AmbiguousAccountError extends AppError {
  constructor(
    message = "Multiple Microsoft connections have access to this account. Please specify connectionId.",
  ) {
    super(message, "account_ambiguous", 409);
    this.name = "AmbiguousAccountError";
  }
}

export class MicrosoftApiError extends AppError {
  constructor(message = "The Microsoft Advertising API request failed.", status = 502, trackingId?: string) {
    super(message, "microsoft_api", status, { microsoftTrackingId: trackingId });
    this.name = "MicrosoftApiError";
  }
}

export class MicrosoftRateLimitError extends AppError {
  constructor(message = "The Microsoft Advertising API rate limit was exceeded. Wait and try again.") {
    super(message, "rate_limited", 429);
    this.name = "MicrosoftRateLimitError";
  }
}

export class ReportError extends AppError {
  constructor(message = "The Microsoft Advertising report failed.", code: AppErrorCode = "report_error") {
    super(message, code, 502);
    this.name = "ReportError";
  }
}

export class ValidationError extends AppError {
  constructor(message: string) {
    super(message, "validation", 400);
    this.name = "ValidationError";
  }
}

export function toHttpStatus(status: number, fallback = 500): number {
  if (Number.isInteger(status) && status >= 200 && status <= 599) {
    return status;
  }
  return fallback;
}

interface MicrosoftLikeError {
  code?: number | string;
  status?: number | string;
  message?: string;
  error?: string;
  error_description?: string;
  trackingId?: string;
}

function asMicrosoftError(error: unknown): MicrosoftLikeError {
  if (error && typeof error === "object") {
    return error as MicrosoftLikeError;
  }
  return { message: error instanceof Error ? error.message : String(error) };
}

function combinedMessage(error: MicrosoftLikeError): string {
  return [error.message, error.error_description, error.error].filter(Boolean).join(" ");
}

export function mapMicrosoftError(error: unknown): AppError {
  if (error instanceof AppError) {
    return error;
  }

  const microsoftError = asMicrosoftError(error);
  const numericCode =
    typeof microsoftError.code === "number"
      ? microsoftError.code
      : typeof microsoftError.status === "number"
        ? microsoftError.status
        : Number.parseInt(String(microsoftError.code ?? microsoftError.status ?? ""), 10);
  const message = combinedMessage(microsoftError);
  const lower = message.toLowerCase();
  const trackingId = microsoftError.trackingId;

  if (
    lower.includes("invalid_grant") ||
    lower.includes("token has been expired or revoked") ||
    lower.includes("invalid refresh token") ||
    lower.includes("aadsts70000") ||
    lower.includes("aadsts700082")
  ) {
    return new ConnectionReauthorizationRequiredError(
      "The Microsoft connection requires reauthorization.",
    );
  }

  if (numericCode === 117 || numericCode === 207 || numericCode === 429 || lower.includes("callrateexceeded") || lower.includes("concurrentrequestoverlimit") || lower.includes("rate limit")) {
    return new MicrosoftRateLimitError();
  }

  if (numericCode === 105 || lower.includes("invalidcredentials") || lower.includes("authentication failed")) {
    return new ConnectionReauthorizationRequiredError(
      "Microsoft Advertising rejected the credentials. Reauthorize this connection.",
    );
  }

  if (numericCode === 401 || lower.includes("unauthenticated") || lower.includes("unauthorized")) {
    return new AuthenticationError(
      "No Microsoft account is connected to this operator, or the access token is no longer valid.",
    );
  }

  if (numericCode === 403 || lower.includes("permission") || lower.includes("not authorized")) {
    return new AuthorizationError(
      "The requested Microsoft Advertising resource is not accessible through the selected connection.",
    );
  }

  if (numericCode === 404 || lower.includes("not found")) {
    return new MicrosoftApiError("The requested Microsoft Advertising resource was not found.", 404, trackingId);
  }

  return new MicrosoftApiError(
    message || "The Microsoft Advertising API request failed.",
    toHttpStatus(numericCode, 502),
    trackingId,
  );
}

export function toToolErrorText(error: unknown): string {
  const mapped = mapMicrosoftError(error);
  return JSON.stringify({
    error: mapped.code,
    message: mapped.message,
    requestId: mapped.microsoftTrackingId,
  });
}
