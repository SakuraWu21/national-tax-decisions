import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  async headers() {
    const alwaysFresh = [
      { key: "Cache-Control", value: "no-store, max-age=0" },
      { key: "CDN-Cache-Control", value: "no-store" },
      { key: "Vercel-CDN-Cache-Control", value: "no-store" },
    ];
    return [
      {
        source: "/data/:path*",
        headers: alwaysFresh,
      },
      {
        source: "/downloads/:path*",
        headers: alwaysFresh,
      },
    ];
  },
};

export default nextConfig;
