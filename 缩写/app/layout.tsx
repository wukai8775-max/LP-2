import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel 报价单生成器",
  description: "根据销售价格表中的产品、公式和格式生成报价 Excel。"
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
