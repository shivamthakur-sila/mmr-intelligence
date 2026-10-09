"use client";

import { useState } from "react";
import type { CuratedDeck, Block } from "@/lib/curate";
import { deckOutline, type OutlineSlide, type Row } from "@/lib/deck-layout";
import { chartPalette, decimalsNeeded, formatNumber, niceAxis } from "@/lib/chart-style";

/**
 * Renders the curated spec as HTML slides at 16:9, mirroring the pptx
 * layout language. This is built from the same spec the .pptx is built
 * from, so it shows the real content and structure — it isn't a
 * pixel-exact render of the file, which would need LibreOffice on the
 * server. Close enough to check the substance before filing.
 */
export default function DeckPreview({
  deck,
  siteName,
  reportMonth,
  coverImage = null,
  initialSlide = 0,
}: {
  deck: CuratedDeck;
  siteName: string;
  reportMonth: string;
  /** The site photo the deck's cover uses, or null for SILA's building. */
  coverImage?: string | null;
  /** The slide shown first; 0, the cover, unless a caller needs another. */
  initialSlide?: number;
}) {
  // The same outline the .pptx is built to and the generate route counts, so
  // the deck the submitter approves is the deck that gets filed - contents and
  // closing slides, and every continuation slide a long section spills onto.
  const slides: OutlineSlide[] = deckOutline(deck.sections);
  const [index, setI] = useState(initialSlide);
  // Clamped: a regenerated deck can be shorter than the slide being viewed.
  const i = Math.min(index, slides.length - 1);
  const current = slides[i];

  return (
    <div>
      <div
        className="relative w-full overflow-hidden rounded-lg"
        style={{ aspectRatio: "16 / 9", background: "#fff", border: "1px solid var(--hairline)" }}
      >
        {current.kind === "cover" && <CoverSlide siteName={siteName} reportMonth={reportMonth} photo={coverImage} />}
        {current.kind === "glance" && <GlanceSlide summary={deck.summary} page={i + 1} />}
        {current.kind === "contents" && <ContentsSlide labels={current.labels} reportMonth={reportMonth} page={i + 1} />}
        {current.kind === "section" && <SectionSlide page={current} pageNumber={i + 1} />}
        {current.kind === "closing" && <ClosingSlide page={i + 1} photo={coverImage} />}
      </div>

      <div className="mt-3 flex items-center justify-between">
        <button
          onClick={() => setI((n) => Math.max(0, n - 1))}
          disabled={i === 0}
          className="text-[13.5px] underline disabled:opacity-30"
          style={{ color: "var(--muted)" }}
        >
          Previous
        </button>
        <span className="text-[13px] tabular-nums" style={{ color: "var(--muted)" }}>
          {i + 1} of {slides.length}
        </span>
        <button
          onClick={() => setI((n) => Math.min(slides.length - 1, n + 1))}
          disabled={i === slides.length - 1}
          className="text-[13.5px] underline disabled:opacity-30"
          style={{ color: "var(--muted)" }}
        >
          Next
        </button>
      </div>

      <div className="mt-4 flex flex-wrap gap-1.5">
        {slides.map((s, n) => (
          <button
            key={n}
            onClick={() => setI(n)}
            title={slideTitle(s)}
            className="h-1.5 rounded-full transition-all"
            style={{
              width: n === i ? 22 : 10,
              background: n === i ? "var(--blue)" : "var(--hairline)",
            }}
          />
        ))}
      </div>
    </div>
  );
}

// Slide geometry in the renderer's inches (13.333 x 7.5), placed by
// percentage so the preview scales with its box. Shared by the chrome, the
// cover and the closing slide, which mirror lib/generate-deck.ts.
const X = (inches: number) => `${(inches / 13.333) * 100}%`;
const Y = (inches: number) => `${(inches / 7.5) * 100}%`;
const box = (x: number, y: number, w: number, h: number): React.CSSProperties => ({
  position: "absolute",
  left: X(x),
  top: Y(y),
  width: X(w),
  height: Y(h),
});
const CHARCOAL = "#3C3C3B";

function Band() {
  return <div style={{ ...box(0, 0, 13.333, 0.75), background: "var(--blue)" }} />;
}

function PageSquare({ n }: { n: number }) {
  return (
    <div
      className="flex items-center justify-center font-semibold"
      style={{ ...box(12.586, 6.753, 0.747, 0.747), background: "var(--sun)", color: CHARCOAL, fontSize: "clamp(6px, 0.9vw, 11px)" }}
    >
      {String(n).padStart(2, "0")}
    </div>
  );
}

function Chrome({ title, page, children }: { title: string; page: number; children: React.ReactNode }) {
  return (
    <div className="absolute inset-0">
      <Band />
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/sila-fm-logo-color.png" alt="SILA" style={{ ...box(11.683, 0.9, 1.4, 0.595), objectFit: "contain" }} />
      <p
        className="flex items-center overflow-hidden whitespace-nowrap"
        style={{ ...box(0.7, 0.92, 10.6, 0.66), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: "clamp(13px, 2.3vw, 30px)" }}
      >
        {title}
      </p>
      <div style={{ ...box(0.7, 1.7, 1.11, 0.04), background: "var(--sun)" }} />
      <div style={{ ...box(13.226, 3.667, 0.107, 0.667), background: "#FFAF00" }} />
      <div className="flex flex-col overflow-hidden" style={box(0.6, 1.95, 12.13, 4.75)}>
        {children}
      </div>
      <PageSquare n={page} />
    </div>
  );
}

/** The cover and closing photo: the site's when it is sharp enough, else SILA's building. */
function PanelPhoto({ photo }: { photo: string | null }) {
  return (
    // eslint-disable-next-line @next/next/no-img-element
    <img src={photo ?? "/bw-building.jpg"} alt="" style={{ ...box(0.76, 0.75, 11.83, 6.0), objectFit: "cover", filter: "grayscale(1)" }} />
  );
}

/** Behind the logo when a site photo is used, as the deck does. */
function LogoPlate({ x, y, w, h }: { x: number; y: number; w: number; h: number }) {
  return <div style={{ ...box(x, y, w, h), background: "#fff" }} />;
}

function CoverSlide({ siteName, reportMonth, photo }: { siteName: string; reportMonth: string; photo: string | null }) {
  return (
    <div className="absolute inset-0">
      <Band />
      <PanelPhoto photo={photo} />
      {photo && <LogoPlate x={5.63} y={0.75} w={2.2} h={1.108} />}
      <div style={{ ...box(0, 0.74, 5.63, 3.01), background: "var(--sun)" }} />
      <div style={{ ...box(0.76, 3.01, 4.87, 0.74), background: "var(--blue)" }} />
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/sila-fm-logo-color.png" alt="SILA" style={{ ...box(5.78, 0.9, 1.9, 0.808), objectFit: "contain" }} />
      <p style={{ ...box(0.76, 1.2, 4.8, 0.5), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: "clamp(9px, 1.4vw, 20px)" }}>
        {prettyMonth(reportMonth)}
      </p>
      <p style={{ ...box(0.76, 2.35, 4.8, 0.4), color: CHARCOAL, letterSpacing: "0.3em", fontSize: "clamp(5px, 0.8vw, 12px)" }}>
        MONTHLY MANAGEMENT REPORT
      </p>
      <p
        className="flex items-center overflow-hidden whitespace-nowrap font-bold"
        style={{ ...box(0.96, 3.01, 2.9, 0.74), fontFamily: "var(--serif)", color: "#fff", letterSpacing: "0.12em", fontSize: "clamp(8px, 1.25vw, 18px)" }}
      >
        {siteName}
      </p>
      <p
        className="flex items-center justify-end underline"
        style={{ ...box(3.85, 3.01, 1.65, 0.74), color: "#fff", fontSize: "clamp(4.5px, 0.65vw, 9.5px)" }}
      >
        www.silagroup.co.in
      </p>
      <PageSquare n={1} />
    </div>
  );
}

function ClosingSlide({ page, photo }: { page: number; photo: string | null }) {
  return (
    <div className="absolute inset-0">
      <Band />
      <PanelPhoto photo={photo} />
      {photo && <LogoPlate x={10.4} y={0.75} w={2.3} h={1.15} />}
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/sila-fm-logo-color.png" alt="SILA" style={{ ...box(10.55, 0.9, 2.0, 0.85), objectFit: "contain" }} />
      <div className="flex items-center" style={{ ...box(0, 4.5, 5.93, 2.27), background: "var(--sun)", paddingLeft: X(0.76) }}>
        <p style={{ fontFamily: "var(--serif)", color: CHARCOAL, letterSpacing: "0.15em", fontSize: "clamp(12px, 2.5vw, 36px)" }}>
          THANK YOU
        </p>
      </div>
      <p
        className="flex items-center justify-center"
        style={{ ...box(0.76, 6.77, 5.17, 0.73), background: "var(--blue)", color: "#fff", fontSize: "clamp(5px, 0.75vw, 10.5px)" }}
      >
        www.silagroup.co.in
      </p>
      <PageSquare n={page} />
    </div>
  );
}

/** Turns "2026-08" into "August 2026", as the deck does. */
function prettyMonth(value: string): string {
  const m = /^(\d{4})-(\d{2})$/.exec(value.trim());
  if (!m) return value;
  const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const name = names[parseInt(m[2], 10) - 1];
  return name ? `${name} ${m[1]}` : value;
}

function GlanceSlide({ summary, page }: { summary: CuratedDeck["summary"]; page: number }) {
  return (
    <Chrome title="This Month at a Glance" page={page}>
      <p
        style={{
          fontFamily: "var(--serif)",
          color: "var(--blue)",
          fontSize: "clamp(9px, 1.5vw, 15px)",
          lineHeight: 1.5,
        }}
      >
        {summary.headline}
      </p>
      <div className="mt-[3%]" style={{ borderTop: "1px solid var(--hairline)" }} />
      <ul className="mt-[3%] space-y-[2.2%]">
        {summary.points.slice(0, 5).map((p, n) => (
          <li key={n} className="flex gap-[2%]">
            <span style={{ width: 3, background: "var(--sun)", flexShrink: 0 }} />
            <span style={{ fontSize: "clamp(7.5px, 1.15vw, 12px)", lineHeight: 1.45 }}>{p}</span>
          </li>
        ))}
      </ul>
    </Chrome>
  );
}

function slideTitle(s: OutlineSlide): string {
  switch (s.kind) {
    case "cover": return "Cover";
    case "glance": return "At a glance";
    case "contents": return "Contents";
    case "closing": return "Thank you";
    case "section": return s.label;
  }
}

/** Mirrors the contents slide in generate-deck.ts: one column, two past eight. */
function ContentsSlide({ labels, reportMonth, page }: { labels: string[]; reportMonth: string; page: number }) {
  const cols = labels.length > 8 ? 2 : 1;
  const perCol = Math.ceil(labels.length / cols);
  const columns = Array.from({ length: cols }, (_, c) => labels.slice(c * perCol, (c + 1) * perCol));
  return (
    <Chrome title={`Contents — ${prettyMonth(reportMonth)}`} page={page}>
      <div className="flex gap-[5%]">
        {columns.map((col, c) => (
          <ul key={c} className="flex-1 space-y-[3%]">
            {col.map((label, n) => (
              <li key={n} className="flex items-center gap-[4%]" style={{ fontSize: "clamp(6.5px, 1.05vw, 12px)" }}>
                <span style={{ width: 4, height: "1.1em", background: "var(--sun)", flexShrink: 0 }} />
                {label}
              </li>
            ))}
          </ul>
        ))}
      </div>
    </Chrome>
  );
}

function SectionSlide({ page, pageNumber }: { page: { label: string; rows: Row[] }; pageNumber: number }) {
  return (
    <Chrome title={page.label} page={pageNumber}>
      <div className="flex h-full flex-col gap-[2.5%]">
        {page.rows.map((row, n) =>
          row.kind === "full" ? (
            <BlockView key={n} block={row.block} />
          ) : (
            // A chart and the table of the same figures, side by side - the
            // arrangement the .pptx uses, so the preview shows the real slide.
            <div key={n} className="flex min-h-0 flex-1 gap-[2.5%]">
              <div className="flex min-h-0 w-1/2 flex-col">
                <BlockView block={row.left} />
              </div>
              <div className="flex min-h-0 w-1/2 flex-col">
                <BlockView block={row.right} />
              </div>
            </div>
          )
        )}
      </div>
    </Chrome>
  );
}

function BlockView({ block }: { block: Block }) {
  const cell: React.CSSProperties = {
    fontSize: "clamp(6px, 0.92vw, 10px)",
    padding: "0.5% 0.9%",
    borderBottom: "1px solid var(--hairline)",
  };

  if (block.type === "kpis") {
    return (
      <div className="flex gap-[1.5%]">
        {block.items.slice(0, 4).map((k, n) => (
          <div
            key={n}
            className="flex-1 px-[2%] py-[1.5%]"
            style={{ background: "#fff", border: "1px solid var(--hairline)", borderLeft: "4px solid var(--sun)" }}
          >
            <p style={{ fontFamily: "var(--serif)", color: "var(--blue)", fontSize: "clamp(12px, 2.3vw, 24px)" }}>
              {k.value}
            </p>
            <p style={{ fontSize: "clamp(5.5px, 0.8vw, 8.5px)", color: "var(--muted)", letterSpacing: "0.08em" }}>
              {k.label.toUpperCase()}
            </p>
          </div>
        ))}
      </div>
    );
  }

  if (block.type === "table") {
    const bands = new Map((block.groups ?? []).map((g) => [g.afterRow, g.label]));
    const out: React.ReactNode[] = [];
    block.rows.forEach((row, ri) => {
      if (bands.has(ri)) {
        out.push(
          <tr key={`band-${ri}`}>
            <td colSpan={block.headers.length} style={{ ...cell, background: "var(--blue)", color: "#fff", fontWeight: 600 }}>
              {bands.get(ri)}
            </td>
          </tr>
        );
      }
      out.push(
        <tr key={`row-${ri}`}>
          {row.map((c, ci) => (
            <td key={ci} style={cell}>
              {c}
            </td>
          ))}
        </tr>
      );
    });

    return (
      <table className="w-full" style={{ borderCollapse: "collapse", tableLayout: "fixed" }}>
        <thead>
          <tr>
            {block.headers.map((h, n) => (
              <th key={n} style={{ ...cell, background: "var(--blue)", color: "#fff", textAlign: "left", fontWeight: 600 }}>
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>{out}</tbody>
      </table>
    );
  }

  if (block.type === "chart") {
    return <ChartView block={block} />;
  }

  if (block.type === "photos") {
    return (
      <div className="grid gap-[1.5%]" style={{ gridTemplateColumns: `repeat(${Math.min(3, block.imageIds.length)}, 1fr)` }}>
        {block.imageIds.slice(0, 6).map((_, n) => (
          <div
            key={n}
            className="flex items-center justify-center"
            style={{ background: "var(--wash)", aspectRatio: "4 / 3", fontSize: "clamp(5.5px, 0.8vw, 9px)", color: "var(--muted)" }}
          >
            photo
          </div>
        ))}
      </div>
    );
  }

  if (block.type === "bullets") {
    return (
      <ul className="space-y-[1.2%]">
        {block.items.map((t, n) => (
          <li key={n} style={{ fontSize: "clamp(6.5px, 1vw, 11px)", lineHeight: 1.45 }}>
            • {t}
          </li>
        ))}
      </ul>
    );
  }

  if (block.type === "note") {
    return (
      <p
        className="px-[1.4%] py-[0.8%]"
        style={{ fontSize: "clamp(5.5px, 0.85vw, 10px)", border: "1px dashed var(--sun)" }}
      >
        {block.text}
      </p>
    );
  }

  return (
    <p style={{ fontSize: "clamp(6.5px, 1vw, 11px)", lineHeight: 1.5 }}>{block.text}</p>
  );
}

// ---------------------------------------------------------------------
// Charts
//
// Drawn as plain SVG rather than with a charting library: the preview only
// has to show the same figures, in the same shape and the same colours, as
// the chart the .pptx will contain. Colours, number formats and axis ranges
// come from lib/chart-style, the module the renderer uses, so the two cannot
// drift again. Layout choices are mirrored by hand: uniform colour for a
// single series, a zero baseline on bars, horizontal bars for long labels.
//
// Elements are keyed by position, never by category or series names: those
// come from the model and can repeat, and a duplicate key makes React drop or
// reuse the wrong element.
// ---------------------------------------------------------------------

const hex = (c: string) => `#${c}`;
const trim = (s: string, max: number) => (s.length > max ? s.slice(0, max - 1) + "…" : s);

function ChartView({ block }: { block: Extract<Block, { type: "chart" }> }) {
  const { chartType, categories, series } = block;
  const dp = decimalsNeeded(series.flatMap((s) => s.values));
  const legend =
    chartType === "pie"
      ? categories.map((c, i) => ({ label: c, color: hex(chartPalette(categories.length)[i]) }))
      : series.length > 1
        ? series.map((s, i) => ({ label: s.name, color: hex(chartPalette(series.length)[i]) }))
        : [];

  // Grows into the space the other blocks leave, but never below roughly the
  // share the renderer guarantees it (1.5in of a 5.07in content area).
  return (
    <div className="flex min-h-0 flex-1 flex-col" style={{ minHeight: "30%" }}>
      {(block.title || block.unit) && (
        <p className="mb-[1%]" style={{ fontFamily: "var(--serif)", fontSize: "clamp(6.5px, 1vw, 11px)" }}>
          {block.title ?? ""}
          {block.unit ? (
            <span style={{ color: "var(--muted)" }}>{block.title ? ` (${block.unit})` : block.unit}</span>
          ) : null}
        </p>
      )}
      <div className="min-h-0 flex-1">
        {chartType === "pie" ? (
          <PieChart categories={categories} values={series[0]?.values ?? []} dp={dp} />
        ) : chartType === "line" ? (
          <LineChart categories={categories} series={series} />
        ) : (
          <BarChart categories={categories} series={series} dp={dp} />
        )}
      </div>
      {legend.length > 0 && (
        <div className="mt-[1%] flex flex-wrap justify-center gap-x-[3%] gap-y-[0.5%]">
          {legend.map((l, n) => (
            <span
              key={n}
              className="flex items-center gap-1"
              style={{ fontSize: "clamp(5px, 0.75vw, 8.5px)", color: "var(--muted)" }}
            >
              <span style={{ width: 6, height: 6, background: l.color, flexShrink: 0 }} />
              {l.label}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

function BarChart({
  categories,
  series,
  dp,
}: {
  categories: string[];
  series: { name: string; values: number[] }[];
  dp: number;
}) {
  if (categories.length === 0 || series.length === 0) return null;
  const longest = Math.max(...categories.map((c) => c.length));
  const horizontal = series.length === 1 && (longest > 14 || categories.length > 8);
  // One colour for a single series; one per series otherwise - as the deck.
  const colors = series.length === 1 ? [hex(chartPalette(1)[0])] : chartPalette(series.length).map(hex);
  // The renderer's axis: it includes zero and reaches below it for negative
  // values, so a negative bar extends the other way instead of vanishing.
  const axis = niceAxis(series.flatMap((s) => s.values), { includeZero: true, pad: 0.1 });
  const span = axis.max - axis.min || 1;

  if (horizontal) {
    const rowH = 100 / categories.length;
    const plotX = 96;
    const plotW = 205;
    const at = (v: number) => plotX + ((v - axis.min) / span) * plotW;
    return (
      <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
        {categories.map((c, i) => {
          const v = series[0].values[i];
          const x0 = at(Math.min(0, v));
          const x1 = at(Math.max(0, v));
          return (
            <g key={i}>
              <text x={plotX - 4} y={i * rowH + rowH / 2} textAnchor="end" dominantBaseline="middle" fontSize={5} fill="#697784">
                {trim(c, 22)}
              </text>
              <rect x={x0} y={i * rowH + rowH * 0.22} width={Math.max(0.5, x1 - x0)} height={rowH * 0.56} fill={colors[0]} />
              <text x={x1 + 3} y={i * rowH + rowH / 2} dominantBaseline="middle" fontSize={5} fill="#2D2D2D">
                {formatNumber(v, dp)}
              </text>
            </g>
          );
        })}
      </svg>
    );
  }

  const top = 14;
  const bottom = 88;
  const at = (v: number) => bottom - ((v - axis.min) / span) * (bottom - top);
  const zero = at(0);
  const colW = 300 / categories.length;
  return (
    <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
      <line x1={18} y1={zero} x2={318} y2={zero} stroke="#C9CDD2" strokeWidth={0.5} />
      {categories.map((c, i) => {
        const groupX = 18 + i * colW;
        const barW = (colW * 0.62) / series.length;
        const offset = (colW - barW * series.length) / 2;
        return (
          <g key={i}>
            {series.map((s, si) => {
              const v = s.values[i];
              const y0 = Math.min(at(v), zero);
              const h = Math.abs(at(v) - zero);
              const x = groupX + offset + si * barW;
              return (
                <g key={si}>
                  <rect x={x} y={y0} width={Math.max(0.5, barW * 0.88)} height={Math.max(0.3, h)} fill={colors[si]} />
                  <text
                    x={x + (barW * 0.88) / 2}
                    y={v >= 0 ? y0 - 2 : y0 + h + 5}
                    textAnchor="middle"
                    fontSize={4.5}
                    fill="#2D2D2D"
                  >
                    {formatNumber(v, dp)}
                  </text>
                </g>
              );
            })}
            <text x={groupX + colW / 2} y={95} textAnchor="middle" fontSize={5} fill="#697784">
              {trim(c, 10)}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

function LineChart({
  categories,
  series,
}: {
  categories: string[];
  series: { name: string; values: number[] }[];
}) {
  if (categories.length === 0 || series.length === 0) return null;
  const colors = chartPalette(series.length).map(hex);
  // The renderer's padded range - a line encodes value by position, so it
  // need not start at zero.
  const axis = niceAxis(series.flatMap((s) => s.values));
  const span = axis.max - axis.min || 1;
  const x = (i: number) => 22 + (i * 292) / Math.max(1, categories.length - 1);
  const y = (v: number) => 82 - ((v - axis.min) / span) * 72;

  return (
    <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
      <line x1={18} y1={82} x2={318} y2={82} stroke="#E5E5E4" strokeWidth={0.5} />
      {series.map((s, si) => (
        <g key={si}>
          <polyline
            points={s.values.map((v, i) => `${x(i)},${y(v)}`).join(" ")}
            fill="none"
            stroke={colors[si]}
            strokeWidth={1.4}
          />
          {s.values.map((v, i) => (
            <circle key={i} cx={x(i)} cy={y(v)} r={1.8} fill={colors[si]} />
          ))}
        </g>
      ))}
      {categories.map((c, i) => (
        <text key={i} x={x(i)} y={92} textAnchor="middle" fontSize={5} fill="#697784">
          {trim(c, 10)}
        </text>
      ))}
    </svg>
  );
}

function PieChart({ categories, values, dp }: { categories: string[]; values: number[]; dp: number }) {
  const colors = chartPalette(categories.length).map(hex);
  const positive = values.map((v) => (Number.isFinite(v) && v > 0 ? v : 0));
  const total = positive.reduce((a, b) => a + b, 0);
  const cx = 50;
  const cy = 50;
  const r = 31;
  if (total <= 0) return null;

  // Each slice's start angle, accumulated up front rather than carried in a
  // variable across the map below - reassigning one during render is what
  // react-hooks/immutability forbids.
  const starts: number[] = [];
  positive.reduce((acc, v) => {
    starts.push(acc);
    return acc + (v / total) * Math.PI * 2;
  }, -Math.PI / 2);

  const arc = (a0: number, a1: number) => {
    const pt = (a: number) => `${cx + r * Math.cos(a)} ${cy + r * Math.sin(a)}`;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    return `M${cx} ${cy} L${pt(a0)} A${r} ${r} 0 ${large} 1 ${pt(a1)} Z`;
  };

  return (
    <svg viewBox="0 0 100 100" className="h-full w-full">
      {positive.map((v, i) => {
        // Zero slices draw nothing. A slice that is the whole pie must be a
        // circle: an arc from an angle back to the same angle has no extent,
        // so a single 100% category used to leave the pie blank.
        if (v <= 0) return null;
        const whole = v / total >= 0.9999;
        const a0 = starts[i];
        const a1 = a0 + (v / total) * Math.PI * 2;
        const mid = (a0 + a1) / 2;
        return (
          <g key={i}>
            {whole ? (
              <circle cx={cx} cy={cy} r={r} fill={colors[i]} />
            ) : (
              <path d={arc(a0, a1)} fill={colors[i]} stroke="#FFFFFF" strokeWidth={0.6} />
            )}
            {v / total > 0.03 && (
              <text
                x={whole ? cx : cx + 36 * Math.cos(mid)}
                y={whole ? cy - r - 4 : cy + 36 * Math.sin(mid)}
                textAnchor="middle"
                dominantBaseline="middle"
                fontSize={4.5}
                fill="#2D2D2D"
              >
                {formatNumber(values[i], dp)}
              </text>
            )}
          </g>
        );
      })}
    </svg>
  );
}
