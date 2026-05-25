/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    serverComponentsExternalPackages: ["exceljs", "xlsx"]
  }
};

export default nextConfig;
