import { escapeHtml, pageHtml } from "@/lib/html";
import type { MicrosoftConnectionRecord, ResumeCandidate } from "@/store/types";

export function connectionsPageHtml(params: {
  connections: MicrosoftConnectionRecord[];
  email?: string | null;
  notice?: string;
  continueAction?: boolean;
}): string {
  const rows = params.connections
    .map(
      (connection) =>
        `<li><code>${escapeHtml(connection.email ?? connection.microsoftSubjectId)}</code> — ${escapeHtml(
          connection.status,
        )}</li>`,
    )
    .join("");

  return pageHtml(
    "Microsoft connections",
    `
      <h1>Microsoft connections</h1>
      ${params.email ? `<p>Operator workspace started as <code>${escapeHtml(params.email)}</code>.</p>` : ""}
      ${params.notice ? `<p class="note">${escapeHtml(params.notice)}</p>` : ""}
      <p>Each Microsoft account is an independent OAuth connection. Advertising customers and accounts are routed to the connection that authorized them.</p>
      <ul>${rows || "<li>No Microsoft accounts connected yet.</li>"}</ul>
      <form method="post" action="/oauth/microsoft/connect">
        <button type="submit">Connect another Microsoft account</button>
      </form>
      ${
        params.continueAction
          ? `<form method="post" action="/oauth/microsoft/connections">
        <button type="submit">Continue to Claude</button>
      </form>`
          : ""
      }
    `,
  );
}

export function resumePageHtml(params: {
  microsoftEmail: string | null;
  candidates: ResumeCandidate[];
  error?: string;
}): string {
  const options = params.candidates
    .map(
      (candidate) =>
        `<option value="${escapeHtml(candidate.operatorId)}">${escapeHtml(
          candidate.primaryEmail ?? candidate.microsoftEmail ?? candidate.operatorId,
        )} — ${candidate.connectionCount} connection(s), created ${escapeHtml(
          candidate.createdAt,
        )}</option>`,
    )
    .join("");

  return pageHtml(
    "Resume or start new",
    `
      <h1>This Microsoft account is already connected</h1>
      <p>Signed in as <code>${escapeHtml(params.microsoftEmail ?? "unknown")}</code>.</p>
      <p>Resume an existing workspace, or start a new isolated operator. Existing Microsoft connections are never merged.</p>
      ${params.error ? `<p class="error">${escapeHtml(params.error)}</p>` : ""}
      <form method="post" action="/oauth/microsoft/resume">
        <label for="choice">What should happen?</label>
        <select id="choice" name="choice" required>
          <option value="resume">Resume an existing workspace</option>
          <option value="start_new">Start a new workspace</option>
        </select>
        <label for="operatorId">Existing workspace</label>
        <select id="operatorId" name="operatorId">
          ${options}
        </select>
        <button type="submit">Continue</button>
      </form>
    `,
  );
}
