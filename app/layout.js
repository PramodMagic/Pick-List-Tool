export const metadata = { title: "Pick List Generator", description: "Stock + Requirement Excel se Pick List banao" };

export default function RootLayout({ children }) {
  return (
    <html lang="hi">
      <body style={{ margin: 0, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif", background: "#f4f6f8", color: "#1b2733" }}>
        {children}
      </body>
    </html>
  );
}
