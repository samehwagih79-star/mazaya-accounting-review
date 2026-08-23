import type { Metadata } from "next";
import "./globals.css";
export const metadata: Metadata = { title: "المساعد المحاسبي | شركة مزايا لتقنية المعلومات", description: "تحليل ومطابقة حسابات العملاء والموردين والبنوك والمخزون لشركة مزايا لتقنية المعلومات" };
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="ar" dir="rtl"><body>{children}</body></html>}
