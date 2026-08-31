import { describe, expect, it } from "vitest";
import { childText, decodeXmlEntities } from "@/microsoft/client/xml";

describe("XML helpers", () => {
  it("decodes report download URLs from SOAP text", () => {
    const xml =
      "<ReportDownloadUrl>https://download.example.com/report.zip?sig=a&amp;file=1</ReportDownloadUrl>";
    expect(childText(xml, "ReportDownloadUrl")).toBe(
      "https://download.example.com/report.zip?sig=a&file=1",
    );
    expect(decodeXmlEntities("a&amp;b")).toBe("a&b");
  });
});
