import { describe, expect, it } from "vitest";
import { createRequestId, getOperatorContext, runWithOperator } from "@/lib/request-context";

describe("request context", () => {
  it("prefers Cloud Trace then request id", () => {
    const req = new Request("http://localhost/mcp", {
      headers: {
        "x-cloud-trace-context": "trace-1/span",
        "x-request-id": "req-1",
      },
    });
    expect(createRequestId(req)).toBe("trace-1");
  });

  it("binds operator context", () => {
    const result = runWithOperator({ requestId: "r1", operatorId: "op-1" }, () => getOperatorContext());
    expect(result.operatorId).toBe("op-1");
    expect(result.requestId).toBe("r1");
  });
});
