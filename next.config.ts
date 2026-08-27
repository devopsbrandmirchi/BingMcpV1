import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  serverExternalPackages: ["@google-cloud/firestore"],
  experimental: {
    cpus: 1,
    webpackBuildWorker: false,
  },
};

export default nextConfig;
