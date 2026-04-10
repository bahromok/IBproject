import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "OfficeAI - Professional AI Agent",
  description: "Your intelligent AI assistant for creating documents, spreadsheets, and automating office tasks.",
  keywords: ["AI", "Agent", "Office", "Documents", "Spreadsheets", "Automation"],
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body>
        {children}
      </body>
    </html>
  );
}
