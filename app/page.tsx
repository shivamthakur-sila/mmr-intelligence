export default function Home() {
  return (
    <main className="min-h-screen">
      <div style={{ height: 4, background: "var(--blue)" }} />
      <div className="mx-auto flex min-h-[70vh] w-full max-w-lg items-center px-6">
        <div>
          <p style={{ fontFamily: "var(--serif)", fontSize: 22, color: "var(--blue)" }}>SILA</p>
          <p style={{ fontSize: 10, letterSpacing: "0.14em", color: "var(--muted)" }}>
            FACILITY MANAGEMENT
          </p>
          <h1 className="mt-8 text-[26px] leading-tight">MMR Intelligence</h1>
          <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
            MMR submissions are made through the link in your reminder email. If you need one,
            email ai@silagroup.co.in.
          </p>
        </div>
      </div>
    </main>
  );
}
