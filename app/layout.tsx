import type { Metadata } from "next";
import { Plus_Jakarta_Sans } from "next/font/google";
import "./globals.css";

// Plus Jakarta Sans across the board — body, display and brand — per the
// Coverdash design system. Loaded through next/font so it's self-hosted rather
// than a render-blocking Google Fonts request.
const jakarta = Plus_Jakarta_Sans({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700"],
  display: "swap",
  variable: "--font-jakarta",
});

export const metadata: Metadata = {
  title: "CalSync · Coverdash",
  description: "Mirror a colleague's recurring meetings onto a new hire.",
  icons: { icon: "/favicon.svg" },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className={jakarta.variable}>
      <body style={{ fontFamily: "var(--font-jakarta), ui-sans-serif, system-ui, sans-serif" }}>
        {children}
      </body>
    </html>
  );
}
