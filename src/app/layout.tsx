import type { Metadata } from "next";
import type { ReactNode } from "react";

export const dynamic = "force-dynamic";

export const metadata: Metadata = {
  title: "Microsoft Advertising MCP Connector",
  description: "Standalone Microsoft Advertising MCP for Claude.ai Custom Connectors",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
