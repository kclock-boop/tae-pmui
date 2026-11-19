import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "품의 · AI 문서 워크스페이스",
  description: "회사 DOCX 양식 그대로, AI로 품의서를 작성하고 편집하세요.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body className="antialiased">{children}</body>
    </html>
  );
}
