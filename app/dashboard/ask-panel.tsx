"use client";

import { useState } from "react";

/**
 * Agent 2's home. Agent 1 is entirely token-driven now — site managers
 * arrive by emailed link and never need an account — so the site list
 * that used to live here has gone. This surface is for asking questions.
 */
export default function AskPanel({ email }: { email: string }) {
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function handleAsk() {
    if (!question.trim()) return;
    setBusy(true);
    setAnswer(null);
    try {
      const res = await fetch(process.env.NEXT_PUBLIC_AGENT2_WEBHOOK_URL!, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ question, email }),
      });
      const data = await res.json();
      setAnswer(data.answer ?? "No answer came back.");
    } catch {
      setAnswer("Couldn't reach Agent 2. Check that its webhook is running.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="min-h-screen">
      <div style={{ height: 4, background: "var(--blue)" }} />
      <div className="mx-auto w-full max-w-2xl px-6 pb-24 pt-14">
        <header className="mb-12 flex items-center justify-between">
          <h1 className="text-[22px]">MMR Intelligence</h1>
          <span className="text-[13px]" style={{ color: "var(--muted)" }}>
            {email}
          </span>
        </header>

        <h2 className="text-[26px] leading-tight">Ask about your sites</h2>
        <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
          Answers come only from filed reports, and always cite where they came from.
        </p>

        <textarea
          rows={3}
          value={question}
          onChange={(e) => setQuestion(e.target.value)}
          placeholder="How did water consumption trend last quarter?"
          className="mt-8"
        />
        <button
          onClick={handleAsk}
          disabled={busy || !question.trim()}
          className="mt-3 rounded-lg px-5 py-2.5 text-[14px] font-medium text-white transition-opacity"
          style={{ background: "var(--blue)", opacity: busy || !question.trim() ? 0.4 : 1 }}
        >
          {busy ? "Looking…" : "Ask"}
        </button>

        {answer && (
          <div
            className="mt-8 rounded-lg px-5 py-4 text-[14.5px] leading-relaxed"
            style={{ background: "var(--paper)", border: "1px solid var(--hairline)" }}
          >
            {answer}
          </div>
        )}
      </div>
    </main>
  );
}
