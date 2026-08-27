import { describe, expect, it } from "vitest";
import { extractMcpToken, parseSseJsonPayload } from "@/mcp/http";

describe("MCP HTTP helpers", () => {
  it("extracts a Bearer token", () => {
    const req = new Request("http://localhost/mcp", {
      headers: { authorization: "Bearer abc.def.ghi" },
    });
    expect(extractMcpToken(req)).toBe("abc.def.ghi");
  });

  it("parses SSE JSON payloads", () => {
    expect(parseSseJsonPayload("data: {\"ok\":true}\n\n")).toBe('{"ok":true}');
  });
});
