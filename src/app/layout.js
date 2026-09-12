import "./globals.css";

export const metadata = {
  title: "Invoice Sync Automation | ERPNext Integration",
  description: "A standalone web portal to parse supplier invoices and auto-generate Purchase Orders, Invoices, and Payments in ERPNext v16.",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
