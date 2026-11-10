import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "文档批量替换助手",
  description: "Word 模板 + Excel 数据，一键生成批量文档",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
