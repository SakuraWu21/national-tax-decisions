import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({ variable: "--font-geist-sans", subsets: ["latin"] });
const geistMono = Geist_Mono({ variable: "--font-geist-mono", subsets: ["latin"] });

export const metadata: Metadata = {
  metadataBase: new URL(process.env.NEXT_PUBLIC_SITE_URL ?? "https://national-tax-decisions.vercel.app"),
  title: {
    default: "全国税务决定书查询平台",
    template: "%s｜全国税务决定书查询平台",
  },
  description: "汇集全国公开的税务处理决定书与税务行政处罚决定书",
  keywords: ["税务处理决定书", "税务行政处罚决定书", "税务稽查", "公开文书"],
  openGraph: {
    title: "全国税务决定书查询平台",
    description: "汇集全国公开的税务处理决定书与税务行政处罚决定书",
    locale: "zh_CN",
    type: "website",
    images: [{ url: "/og.png", width: 1733, height: 909, alt: "全国税务决定书查询平台数据看板" }],
  },
  twitter: {
    card: "summary_large_image",
    title: "全国税务决定书查询平台",
    description: "汇集全国公开的税务处理决定书与税务行政处罚决定书",
    images: ["/og.png"],
  },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN" className={`${geistSans.variable} ${geistMono.variable}`}>
      <body>{children}</body>
    </html>
  );
}
