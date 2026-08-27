import { randomBytes } from "node:crypto";
import { encryptRefreshToken } from "@/lib/crypto";
import { AppError } from "@/lib/errors";
import { escapeHtml, htmlResponse, pageHtml } from "@/lib/html";
import { logger } from "@/lib/logger";
import {
  clearOAuthCookieHeaders,
  operatorCookieHeader,
  readOAuthCookies,
  requestIsHttps,
  resumeCookieHeader,
} from "@/lib/oauth-cookies";
import { createMcpSessionId } from "@/mcp/oauth/tokens";
import { exchangeAuthorizationCode, verifyOAuthState } from "@/microsoft/auth/oauth";
import { bindMicrosoftIdentityToOperator, createOperatorWithConnection } from "@/oauth/bind-connection";
import { connectionsPageHtml, resumePageHtml } from "@/oauth/pages";
import { getAppStore } from "@/store/app-store";
import type { ResumeCandidate } from "@/store/types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function appendCookies(headers: Headers, cookies: string[]): void {
  for (const cookie of cookies) {
    headers.append("Set-Cookie", cookie);
  }
}

export async function GET(req: Request) {
  const url = new URL(req.url);
  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");
  const oauthError = url.searchParams.get("error");
  const forgedOperatorId = url.searchParams.get("operatorId");
  const secure = requestIsHttps(req);
  const cookies = readOAuthCookies(req);

  if (oauthError) {
    const deniedHeaders = new Headers();
    appendCookies(deniedHeaders, clearOAuthCookieHeaders(secure));
    return htmlResponse(
      pageHtml(
        "Microsoft authorization denied",
        `<h1>Microsoft authorization denied</h1><p class="error">${escapeHtml(
          oauthError,
        )}</p><p>Return to Claude and connect the connector again.</p>`,
      ),
      400,
      deniedHeaders,
    );
  }

  try {
    if (forgedOperatorId) {
      throw new AppError("Operator identity cannot be supplied by the browser.", "unauthorized", 401);
    }
    if (!code || !state || !cookies.state || !cookies.verifier) {
      throw new AppError("The Microsoft authorization response is incomplete.", "session_invalid", 400);
    }
    if (state !== cookies.state) {
      throw new AppError("OAuth state did not match.", "unauthorized", 401);
    }
    const signed = verifyOAuthState(state);
    const transaction = await getAppStore().consumeTransaction(signed.nonce);
    if (transaction.microsoftState && transaction.microsoftState !== state) {
      throw new AppError("OAuth state did not match.", "unauthorized", 401);
    }
    if (transaction.codeVerifier && transaction.codeVerifier !== cookies.verifier) {
      throw new AppError("OAuth PKCE verifier did not match.", "unauthorized", 401);
    }

    const identity = await exchangeAuthorizationCode({
      code,
      codeVerifier: transaction.codeVerifier || cookies.verifier,
    });

    if (transaction.operation === "add_connection" || transaction.operation === "reconnect") {
      if (!transaction.operatorId) {
        throw new AppError("The Microsoft connection was not bound to an operator.", "unauthorized", 401);
      }
      const { operator, connection } = await bindMicrosoftIdentityToOperator({
        operatorId: transaction.operatorId,
        identity,
      });
      logger.info("Flow 2 Microsoft connection bound", {
        operatorId: operator.operatorId,
        connectionId: connection.connectionId,
      });
      const sessionId = cookies.operator?.sessionId || createMcpSessionId();
      const connections = await getAppStore().listConnections(operator.operatorId);
      const headers = new Headers();
      headers.append(
        "Set-Cookie",
        operatorCookieHeader({ operatorId: operator.operatorId, sessionId, secure }),
      );
      return htmlResponse(
        connectionsPageHtml({
          connections,
          email: identity.email,
          notice: "Microsoft account connected.",
          continueAction: Boolean(cookies.pending || transaction.mcpPending),
        }),
        200,
        headers,
      );
    }

    if (transaction.operation !== "bootstrap_operator") {
      throw new AppError("OAuth transaction is not valid for this callback.", "unauthorized", 401);
    }

    const existing = (await getAppStore().listConnectionsByMicrosoftSubject(identity.microsoftSubjectId))
      .filter((item) => item.status !== "disconnected");

    if (existing.length > 0) {
      const candidates: ResumeCandidate[] = [];
      for (const connection of existing) {
        const operator = await getAppStore().getOperator(connection.operatorId);
        if (!operator) {
          continue;
        }
        const count = (await getAppStore().listConnections(operator.operatorId)).length;
        candidates.push({
          operatorId: operator.operatorId,
          primaryEmail: operator.primaryEmail,
          microsoftEmail: connection.email,
          connectionCount: count,
          createdAt: operator.createdAt,
        });
      }
      const resumeNonce = randomBytes(16).toString("hex");
      await getAppStore().createTransaction({
        nonce: resumeNonce,
        operatorId: null,
        operation: "resume_choice",
        connectionId: null,
        mcpPending: transaction.mcpPending ?? cookies.pending ?? null,
        codeVerifier: null,
        microsoftState: null,
        exp: Date.now() + 15 * 60 * 1000,
        resumeCandidates: candidates,
        pendingMicrosoftSubjectId: identity.microsoftSubjectId,
        pendingMicrosoftEmail: identity.email,
        pendingMicrosoftDisplayName: identity.displayName,
        pendingRefreshTokenEncrypted: identity.refreshToken
          ? encryptRefreshToken(identity.refreshToken)
          : null,
        pendingMicrosoftScopes: identity.scopes,
      });
      const headers = new Headers();
      headers.append("Set-Cookie", resumeCookieHeader(resumeNonce, secure));
      return htmlResponse(
        resumePageHtml({
          microsoftEmail: identity.email,
          candidates,
        }),
        200,
        headers,
      );
    }

    const created = await createOperatorWithConnection(identity);
    logger.info("Created operator from Microsoft bootstrap", {
      operatorId: created.operator.operatorId,
      connectionId: created.connection.connectionId,
    });
    const sessionId = createMcpSessionId();
    const connections = await getAppStore().listConnections(created.operator.operatorId);
    const headers = new Headers();
    headers.append(
      "Set-Cookie",
      operatorCookieHeader({
        operatorId: created.operator.operatorId,
        sessionId,
        secure,
      }),
    );
    return htmlResponse(
      connectionsPageHtml({
        connections,
        email: identity.email,
        notice: "Microsoft account connected.",
        continueAction: Boolean(transaction.mcpPending || cookies.pending),
      }),
      200,
      headers,
    );
  } catch (error) {
    logger.error("Microsoft OAuth callback failed", {
      errorMessage: error instanceof Error ? error.message : "unknown",
    });
    const headers = new Headers();
    appendCookies(headers, clearOAuthCookieHeaders(secure));
    return htmlResponse(
      pageHtml(
        "Microsoft authorization failed",
        `<h1>Microsoft authorization failed</h1><p class="error">${escapeHtml(
          error instanceof Error ? error.message : "Authorization failed",
        )}</p>`,
      ),
      400,
      headers,
    );
  }
}
