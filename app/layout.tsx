import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "MMR Intelligence",
  description: "MMR tracking and site insights",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col font-sans">{children}</body>
    </html>
  );
}
