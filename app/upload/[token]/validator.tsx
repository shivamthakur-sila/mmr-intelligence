"use client";

import { useEffect, useState } from "react";
import DeckPreview from "./deck-preview";
import type { CuratedDeck } from "@/lib/curate";
import { uploadViaStorage } from "@/lib/upload-via-storage";

type Section = {
  key: string;
  label: string;
  askType: "upload" | "text" | "photos";
  present: boolean;
  request?: string;
};

type Step = "upload" | "review" | "preview" | "filed";

const ASK_HINT: Record<Section["askType"], string> = {
  upload: "Attach the sheet, or type what you have",
  text: "Type what you have",
  photos: "Describe it, or attach photos",
};

export default function Validator({
  token,
  siteName,
  reportMonth,
}: {
  token: string;
  siteName: string;
  reportMonth: string;
}) {
  const [step, setStep] = useState<Step>("upload");
  const [files, setFiles] = useState<File[]>([]);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [rejection, setRejection] = useState<{ reason: string; evidence: string } | null>(null);
  const [uploadProgress, setUploadProgress] = useState<{ done: number; total: number } | null>(null);
  // Set when the slides name this month but the dated figures underneath are
  // from another period. The submitter decides, not the system: refusing a
  // whole submission over one stale embedded sheet is too harsh, and carrying
  // on silently is how last year's figures reach a client under this month.
  const [periodConflict, setPeriodConflict] = useState<string | null>(null);
  const [periodConfirmed, setPeriodConfirmed] = useState(false);

  const [sessionId, setSessionId] = useState("");
  const [sections, setSections] = useState<Section[]>([]);
  const [photoCount, setPhotoCount] = useState(0);
  const [revealed, setRevealed] = useState(0);

  const [answers, setAnswers] = useState<Record<string, string>>({});
  const [skipped, setSkipped] = useState<Record<string, boolean>>({});
  const [attachments, setAttachments] = useState<Record<string, File>>({});

  const [deck, setDeck] = useState<CuratedDeck | null>(null);
  const [slideCount, setSlideCount] = useState(0);
  const [coverPreview, setCoverPreview] = useState<string | null>(null);
  const [storageWarning, setStorageWarning] = useState<string | null>(null);
  const [tokenWarning, setTokenWarning] = useState<string | null>(null);

  useEffect(() => {
    if (step !== "review" || sections.length === 0 || revealed >= sections.length) return;
    const t = setTimeout(() => setRevealed((n) => n + 1), 55);
    return () => clearTimeout(t);
  }, [step, sections, revealed]);

  function addFiles(list: FileList | File[]) {
    setFiles((prev) => [...prev, ...Array.from(list).filter((f) => /\.(pptx|xlsx)$/i.test(f.name))]);
    setError("");
    setRejection(null);
  }

  async function runReview() {
    if (files.length === 0) return;
    setBusy(true);
    setError("");
    setRejection(null);
    setUploadProgress({ done: 0, total: files.length });

    try {
      // Straight to storage first - only the paths go through Vercel.
      const refs = await uploadViaStorage(token, files, (done, total) =>
        setUploadProgress({ done, total })
      );
      setUploadProgress(null);

      const res = await fetch(`/api/upload/${token}/review`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ files: refs }),
      });
      const data = await res.json();

      if (res.status === 422 && data.rejected) {
        setRejection({ reason: data.reason, evidence: data.evidence });
        setBusy(false);
        return;
      }
      if (!res.ok) {
        setError(data.error || "The submission couldn't be checked.");
        setBusy(false);
        return;
      }

      setSessionId(data.sessionId);
      setPeriodConflict(data.validation?.verdict === "period_conflict" ? data.validation.evidence : null);
      setPeriodConfirmed(false);
      setSections(data.sections);
      setPhotoCount(data.photoCount);
      setRevealed(0);
      setStep("review");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't reach the server.");
      setUploadProgress(null);
    } finally {
      setBusy(false);
    }
  }

  async function runGenerate() {
    setBusy(true);
    setError("");

    try {
      const attachedFiles = Object.values(attachments);
      const attachmentRefs =
        attachedFiles.length > 0 ? await uploadViaStorage(token, attachedFiles) : [];

      const res = await fetch(`/api/upload/${token}/generate`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          sessionId,
          answers: sections
            .filter((s) => !s.present && !skipped[s.key] && (answers[s.key] ?? "").trim())
            .map((s) => ({ label: s.label, answer: answers[s.key].trim() })),
          attachments: attachmentRefs,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "The deck couldn't be built.");
        setBusy(false);
        return;
      }
      setDeck(data.deck);
      setSlideCount(data.slideCount);
      setCoverPreview(data.coverPreview ?? null);
      setStep("preview");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't reach the server.");
    } finally {
      setBusy(false);
    }
  }

  async function runCommit() {
    setBusy(true);
    setError("");
    try {
      const res = await fetch(`/api/upload/${token}/commit`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ sessionId }),
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || "Couldn't file it.");
        setBusy(false);
        return;
      }
      setStorageWarning(data.storageWarning ?? null);
      setTokenWarning(data.tokenWarning ?? null);
      setStep("filed");
    } catch {
      setError("Couldn't reach the server.");
    } finally {
      setBusy(false);
    }
  }

  const missing = sections.filter((s) => !s.present);
  const found = sections.filter((s) => s.present);
  const open = missing.filter(
    (s) => !skipped[s.key] && !(answers[s.key] ?? "").trim() && !attachments[s.key]
  );

  return (
    <main className="min-h-screen">
      <div style={{ height: 4, background: "var(--blue)" }} />
      <div className="mx-auto w-full max-w-2xl px-6 pb-24 pt-14">
        <header className="mb-12 flex items-center justify-between">
          <div>
            <p style={{ fontFamily: "var(--serif)", fontSize: 19, color: "var(--blue)", lineHeight: 1.1 }}>
              SILA
            </p>
            <p style={{ fontSize: 9.5, letterSpacing: "0.14em", color: "var(--muted)" }}>
              FACILITY MANAGEMENT
            </p>
          </div>
          <div className="text-right">
            <p className="text-[14px]">{siteName}</p>
            <p className="text-[13px]" style={{ color: "var(--muted)" }}>
              {reportMonth}
            </p>
          </div>
        </header>

        {step === "upload" && (
          <section className="step-in">
            <h1 className="text-[29px] leading-tight">Submit this month&apos;s MMR</h1>
            <p className="mt-3 max-w-lg text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
              Add whatever you have — a draft deck, working spreadsheets, or both. Everything is checked
              against what this site is required to report before anything is built.
            </p>

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragging(false);
                if (e.dataTransfer.files?.length) addFiles(e.dataTransfer.files);
              }}
              className="mt-9 rounded-xl px-6 py-12 text-center transition-colors"
              style={{
                background: dragging ? "rgba(247,163,40,0.06)" : "var(--paper)",
                border: `1px dashed ${dragging ? "var(--sun)" : "var(--hairline)"}`,
              }}
            >
              <p className="text-[15px]">Drop files here</p>
              <p className="mt-1 text-[13px]" style={{ color: "var(--muted)" }}>
                PowerPoint or Excel, as many as you have
              </p>
              <label
                tabIndex={0}
                className="mt-5 inline-block rounded-lg px-4 py-2 text-[14px] font-medium text-white"
                style={{ background: "var(--blue)" }}
              >
                Choose files
                <input
                  type="file"
                  accept=".pptx,.xlsx"
                  multiple
                  className="hidden"
                  onChange={(e) => e.target.files && addFiles(e.target.files)}
                />
              </label>
            </div>

            {files.length > 0 && (
              <ul className="mt-6" style={{ borderTop: "1px solid var(--hairline)" }}>
                {files.map((f, i) => (
                  <li
                    key={`${f.name}-${i}`}
                    className="flex items-center justify-between py-3"
                    style={{ borderBottom: "1px solid var(--hairline)" }}
                  >
                    <span className="truncate pr-4 text-[14px]">{f.name}</span>
                    <div className="flex shrink-0 items-center gap-4">
                      <span className="text-[13px] tabular-nums" style={{ color: "var(--muted)" }}>
                        {(f.size / 1024 / 1024).toFixed(1)} MB
                      </span>
                      <button
                        onClick={() => setFiles((p) => p.filter((_, j) => j !== i))}
                        className="text-[13px] underline"
                        style={{ color: "var(--muted)" }}
                      >
                        Remove
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
            )}

            {rejection && (
              <div
                className="mt-6 rounded-lg px-4 py-4"
                style={{ background: "rgba(199,62,62,0.06)" }}
              >
                <p className="text-[14px] font-medium" style={{ color: "#a33b3b" }}>
                  {rejection.reason}
                </p>
                <p className="mt-2 text-[13.5px] leading-relaxed" style={{ color: "#a33b3b" }}>
                  {rejection.evidence}
                </p>
                <p className="mt-3 text-[13px]" style={{ color: "var(--muted)" }}>
                  Remove the wrong file and add the right one, then check again.
                </p>
              </div>
            )}

            {error && <ErrorNote text={error} />}

            {busy ? (
              <Working
                label={
                  uploadProgress
                    ? `Uploading ${uploadProgress.done} of ${uploadProgress.total} file${uploadProgress.total === 1 ? "" : "s"}…`
                    : "Reading the files and checking them against this site's required parameters…"
                }
                note={
                  uploadProgress
                    ? "Large decks take a moment to upload."
                    : "This usually takes around half a minute. You can leave this tab open."
                }
              />
            ) : (
              <button
                onClick={runReview}
                disabled={files.length === 0}
                className="mt-8 w-full rounded-lg py-3 text-[15px] font-medium text-white transition-opacity"
                style={{ background: "var(--blue)", opacity: files.length === 0 ? 0.4 : 1 }}
              >
                Check the submission
              </button>
            )}
          </section>
        )}

        {step === "review" && (
          <section className="step-in">
            <h1 className="text-[29px] leading-tight">
              {missing.length === 0
                ? "Everything\u2019s here"
                : `${missing.length} ${missing.length === 1 ? "gap" : "gaps"} to settle`}
            </h1>
            <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
              {found.length} of {sections.length} required parameters have real content
              {photoCount > 0 && `, and ${photoCount} site photos were found`}. Fill a gap or skip it —
              anything skipped is simply left out of the deck rather than shown as missing.
            </p>

            <ul className="mt-9">
              {sections.slice(0, revealed).map((s, i) => {
                const resolved = !!skipped[s.key] || !!(answers[s.key] ?? "").trim() || !!attachments[s.key];
                return (
                  <li
                    key={s.key}
                    className="settle py-4"
                    style={{
                      borderTop: i === 0 ? "1px solid var(--hairline)" : undefined,
                      borderBottom: "1px solid var(--hairline)",
                    }}
                  >
                    <div className="flex items-start gap-3">
                      <span
                        aria-label={s.present ? "Found" : resolved ? "Settled" : "Missing"}
                        className="mt-1.5 block h-[7px] w-[7px] shrink-0 rounded-full"
                        style={{
                          background: s.present
                            ? "var(--blue)"
                            : resolved
                              ? "var(--hairline)"
                              : "var(--sun)",
                        }}
                      />
                      <div className="min-w-0 flex-1">
                        <p className="text-[15px]">{s.label}</p>

                        {!s.present && (
                          <>
                            <p
                              className="mt-1 text-[13.5px] leading-relaxed"
                              style={{ color: "var(--muted)" }}
                            >
                              {s.request || "Not found in the files provided."}
                            </p>

                            {skipped[s.key] ? (
                              <p className="mt-2 text-[13px]" style={{ color: "var(--muted)" }}>
                                Skipped — left out of the deck.{" "}
                                <button
                                  onClick={() => setSkipped((p) => ({ ...p, [s.key]: false }))}
                                  className="underline"
                                >
                                  Undo
                                </button>
                              </p>
                            ) : (
                              <div className="mt-3">
                                <textarea
                                  rows={2}
                                  value={answers[s.key] ?? ""}
                                  onChange={(e) =>
                                    setAnswers((p) => ({ ...p, [s.key]: e.target.value }))
                                  }
                                  placeholder={ASK_HINT[s.askType]}
                                />
                                <div className="mt-2 flex flex-wrap items-center gap-4">
                                  <label
                                    tabIndex={0}
                                    className="text-[13px] underline"
                                    style={{ color: "var(--blue)" }}
                                  >
                                    {attachments[s.key]
                                      ? `Attached: ${attachments[s.key].name}`
                                      : "Attach a file"}
                                    <input
                                      type="file"
                                      accept=".xlsx,.jpg,.jpeg,.png"
                                      className="hidden"
                                      onChange={(e) => {
                                        const f = e.target.files?.[0];
                                        if (f) setAttachments((p) => ({ ...p, [s.key]: f }));
                                      }}
                                    />
                                  </label>
                                  <button
                                    onClick={() => setSkipped((p) => ({ ...p, [s.key]: true }))}
                                    className="text-[13px] underline"
                                    style={{ color: "var(--muted)" }}
                                  >
                                    Skip this
                                  </button>
                                </div>
                              </div>
                            )}
                          </>
                        )}
                      </div>
                    </div>
                  </li>
                );
              })}
            </ul>

            {revealed >= sections.length && (
              <>
                {error && <ErrorNote text={error} />}
                {open.length > 0 && (
                  <p className="mt-6 text-[13.5px]" style={{ color: "var(--muted)" }}>
                    {open.length} {open.length === 1 ? "gap is" : "gaps are"} still open. Build now and
                    they&apos;re left out.
                  </p>
                )}
                {periodConflict && (
                  <div
                    className="mt-6 rounded-lg px-4 py-3 text-[13.5px] leading-relaxed"
                    style={{ border: "1px solid var(--sun)", background: "var(--paper)" }}
                  >
                    <p className="font-medium">Some figures may be from a different period</p>
                    <p className="mt-1" style={{ color: "var(--muted)" }}>
                      {periodConflict}
                    </p>
                    <p className="mt-2" style={{ color: "var(--muted)" }}>
                      If a workbook was carried over from an earlier report, upload the current one
                      instead. If these figures are right, confirm below — each one will be shown
                      against the period its own row gives, never relabelled as {reportMonth}.
                    </p>
                    <label className="mt-3 flex items-center gap-2">
                      <input
                        type="checkbox"
                        checked={periodConfirmed}
                        onChange={(e) => setPeriodConfirmed(e.target.checked)}
                      />
                      These figures are correct as submitted
                    </label>
                  </div>
                )}
                {busy ? (
                  <Working
                    label="Curating the content and building the deck…"
                    note="This is the slowest step — usually a minute or two, since every section is written from your data and the photos are placed individually."
                  />
                ) : (
                  <button
                    onClick={runGenerate}
                    disabled={!!periodConflict && !periodConfirmed}
                    className="mt-6 w-full rounded-lg py-3 text-[15px] font-medium text-white disabled:opacity-40"
                    style={{ background: "var(--sun)" }}
                  >
                    Build the deck
                  </button>
                )}
              </>
            )}
          </section>
        )}

        {step === "preview" && deck && (
          <section className="step-in">
            <h1 className="text-[29px] leading-tight">Check it over</h1>
            <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
              {slideCount} slides, built from {deck.sections.length} sections that had real content.
              Every figure was checked against what you submitted.
            </p>

            {(deck.submitterWarnings ?? []).length > 0 && (
              <div
                className="mt-5 rounded-lg px-4 py-3 text-[13.5px] leading-relaxed"
                style={{ background: "var(--wash)", borderLeft: "3px solid var(--sun)" }}
              >
                <p className="font-medium">Before you file, about your submission</p>
                <ul className="mt-2 list-disc space-y-1 pl-5" style={{ color: "var(--muted)" }}>
                  {(deck.submitterWarnings ?? []).map((w, n) => (
                    <li key={n}>{w}</li>
                  ))}
                </ul>
                <p className="mt-2" style={{ color: "var(--muted)" }}>
                  This is for you only; it is not in the deck.
                </p>
              </div>
            )}

            {(() => {
              const removed = (deck.grounding ?? []).filter((g) => g.action === "dropped");
              if (removed.length === 0) return null;
              return (
                <details
                  className="mt-5 rounded-lg px-4 py-3 text-[13.5px] leading-relaxed"
                  style={{ background: "var(--wash)" }}
                >
                  <summary className="cursor-pointer">
                    {removed.length} {removed.length === 1 ? "item was" : "items were"} left out because
                    {removed.length === 1 ? " its figures aren't" : " their figures aren't"} in your files
                  </summary>
                  <ul className="mt-2 space-y-1" style={{ color: "var(--muted)" }}>
                    {removed.map((g, n) => (
                      <li key={n}>
                        {g.where}: {g.value.length > 140 ? g.value.slice(0, 139) + "…" : g.value}
                      </li>
                    ))}
                  </ul>
                  <p className="mt-2" style={{ color: "var(--muted)" }}>
                    If a figure should be there, go back and add it as an answer to the gap.
                  </p>
                </details>
              );
            })()}

            <div className="mt-8">
              <DeckPreview deck={deck} siteName={siteName} reportMonth={reportMonth} coverImage={coverPreview} />
            </div>

            {error && <ErrorNote text={error} />}

            <div className="mt-8 flex flex-col gap-3 sm:flex-row">
              <a
                href={`/api/upload/${token}/download?session=${sessionId}`}
                className="flex-1 rounded-lg py-3 text-center text-[15px] font-medium"
                style={{ border: "1px solid var(--hairline)", background: "var(--paper)" }}
              >
                Download to review
              </a>
              <button
                onClick={runCommit}
                disabled={busy}
                className="flex-1 rounded-lg py-3 text-[15px] font-medium text-white transition-opacity"
                style={{ background: "var(--blue)", opacity: busy ? 0.5 : 1 }}
              >
                {busy ? "Filing…" : "Submit and file it"}
              </button>
            </div>
            <button
              onClick={() => setStep("review")}
              className="mt-4 w-full text-center text-[13.5px] underline"
              style={{ color: "var(--muted)" }}
            >
              Back to the gaps
            </button>
          </section>
        )}

        {step === "filed" && (
          <section className="step-in">
            <h1 className="text-[29px] leading-tight">Filed</h1>
            <p className="mt-3 text-[15px] leading-relaxed" style={{ color: "var(--muted)" }}>
              {reportMonth}&apos;s MMR for {siteName} is recorded as complete. You won&apos;t get any
              further reminders about it.
            </p>

            {storageWarning && (
              <p className="mt-6 rounded-lg px-4 py-3 text-[13px] leading-relaxed" style={{ background: "var(--wash)", color: "var(--muted)" }}>
                The deck was registered, but couldn&apos;t be saved to storage: {storageWarning}. The
                record is filed; someone may need to re-upload the file itself.
              </p>
            )}

            {tokenWarning && (
              <p className="mt-4 rounded-lg px-4 py-3 text-[13px] leading-relaxed" style={{ background: "var(--wash)", color: "var(--muted)" }}>
                {tokenWarning}. Your submission is filed; this link will still refuse a second one.
              </p>
            )}

            {/* No download here. Once a month is filed this link is closed -
                resolveToken refuses it, by design, for every route including
                download - so a button here could only ever return an error.
                A copy can be taken with "Download to review" before filing. */}
          </section>
        )}
      </div>
    </main>
  );
}


/**
 * Shown while the server is busy. The bar is indeterminate on purpose:
 * extraction, site/month validation and the checklist review all happen
 * inside one request, so any percentage would be made up. The elapsed
 * counter is what actually reassures someone it hasn't frozen - the
 * review takes around half a minute and generation longer.
 */
function Working({ label, note }: { label: string; note?: string }) {
  const [seconds, setSeconds] = useState(0);
  useEffect(() => {
    const t = setInterval(() => setSeconds((n) => n + 1), 1000);
    return () => clearInterval(t);
  }, []);

  return (
    <div
      className="mt-8 rounded-lg px-5 py-5"
      style={{ background: "var(--paper)", border: "1px solid var(--hairline)" }}
      role="status"
      aria-live="polite"
    >
      <div className="flex items-baseline justify-between gap-4">
        <p className="text-[14.5px]">{label}</p>
        <span className="shrink-0 text-[13px] tabular-nums" style={{ color: "var(--muted)" }}>
          {seconds}s
        </span>
      </div>
      <div className="working-track mt-3">
        <div className="working-fill" />
      </div>
      {note && (
        <p className="mt-3 text-[13px] leading-relaxed" style={{ color: "var(--muted)" }}>
          {note}
        </p>
      )}
    </div>
  );
}

function ErrorNote({ text }: { text: string }) {
  return (
    <p
      className="mt-6 rounded-lg px-4 py-3 text-[13.5px] leading-relaxed"
      style={{ background: "rgba(199,62,62,0.06)", color: "#a33b3b" }}
    >
      {text}
    </p>
  );
}