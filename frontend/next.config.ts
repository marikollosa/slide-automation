import type { NextConfig } from "next";

const repo = "slide-automation";

const nextConfig: NextConfig = {
  // Required for GitHub Pages static hosting
  output: "export",

  // GitHub Pages serves at /REPO_NAME/, not /
  basePath: process.env.NODE_ENV === "production" ? `/${repo}` : "",
  assetPrefix: process.env.NODE_ENV === "production" ? `/${repo}/` : "",

  // GitHub Pages does not support Next Image optimization
  images: {
    unoptimized: true,
  },
};

export default nextConfig;

