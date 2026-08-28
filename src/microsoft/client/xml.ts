export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function stripXmlNamespaces(xml: string): string {
  return xml
    .replace(/<\/([A-Za-z_][\w.-]*):/g, "</")
    .replace(/<([A-Za-z_][\w.-]*):/g, "<")
    .replace(/\sxmlns(?::\w+)?="[^"]*"/g, "")
    .replace(/\s(?:xsi|i):nil="[^"]*"/g, "");
}

export function childText(xml: string, tag: string): string | null {
  const match = xml.match(new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "i"));
  if (!match) {
    return null;
  }
  const value = (match[1] ?? "").trim();
  return value.length > 0 ? value : null;
}

export function childBlocks(xml: string, tag: string): string[] {
  const regex = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}>`, "gi");
  const blocks: string[] = [];
  let match = regex.exec(xml);
  while (match) {
    blocks.push(match[1] ?? "");
    match = regex.exec(xml);
  }
  return blocks;
}

export function longArrayXml(tag: string, values: Array<string | number>): string {
  const items = values.map((value) => `<a1:long>${xmlEscape(String(value))}</a1:long>`).join("");
  return `<${tag} xmlns:a1="http://schemas.microsoft.com/2003/10/Serialization/Arrays">${items}</${tag}>`;
}

export function stringArrayXml(tag: string, itemTag: string, values: string[]): string {
  const items = values.map((value) => `<${itemTag}>${xmlEscape(value)}</${itemTag}>`).join("");
  return `<${tag}>${items}</${tag}>`;
}
