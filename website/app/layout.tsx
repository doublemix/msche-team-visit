import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import Link from "next/link";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "MSCHE Team Visit",
  description: "MSCHE Team Visit",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body
        className={`${geistSans.variable} ${geistMono.variable} min-h-[100vh] antialiased`}
      >
        <div className="min-h-[100vh] bg-gray-500 pt-16">
          <div className="mx-auto max-w-[640px]">
            <div className="mb-[-20px] flex flex-row gap-2 rounded-t-2xl bg-gray-400 px-8 pt-4 pb-8">
              <Link
                className="rounded-md bg-gray-300 px-2 py-1 text-blue-700 hover:bg-gray-200 active:bg-gray-200 active:inset-shadow-xs active:inset-shadow-gray-500"
                href="/"
              >
                Home
              </Link>
              <Link
                className="rounded-md bg-gray-300 px-2 py-1 text-blue-700 hover:bg-gray-200 active:bg-gray-200 active:inset-shadow-xs active:inset-shadow-gray-500"
                href="/meetings"
              >
                All Meetings
              </Link>
            </div>
          </div>
          <div className="mx-auto min-h-[640px] max-w-[640px] rounded-2xl bg-white p-8">
            {children}
          </div>
        </div>
      </body>
    </html>
  );
}
