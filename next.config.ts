import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  typescript: {
    ignoreBuildErrors: true,
  },
  reactStrictMode: false,
  webpack: (config, { isServer }) => {
    if (isServer) {
      // Don't bundle native addons - load them at runtime
      config.externals.push({
        'canvas': 'commonjs canvas',
        'chartjs-node-canvas': 'commonjs chartjs-node-canvas',
      });
    }
    return config;
  },
};

export default nextConfig;
