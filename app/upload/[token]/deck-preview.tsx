"use client";

import { useState } from "react";
import type { CuratedDeck, Block } from "@/lib/curate";
import { sectionSlides, type Row } from "@/lib/deck-layout";

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
}: {
  deck: CuratedDeck;
  siteName: string;
  reportMonth: string;
}) {
  // Sections are split here exactly as the .pptx splits them, using the same
  // measurements, so the slide the submitter approves is the slide that gets
  // filed — including the continuation slides a long section spills onto.
  const slides = [
    { kind: "cover" as const },
    { kind: "glance" as const },
    ...deck.sections.flatMap((s) =>
      sectionSlides(s).map((page) => ({ kind: "section" as const, page }))
    ),
  ];
  const [i, setI] = useState(0);
  const current = slides[i];

  return (
    <div>
      <div
        className="relative w-full overflow-hidden rounded-lg"
        style={{ aspectRatio: "16 / 9", background: "#fff", border: "1px solid var(--hairline)" }}
      >
        {current.kind === "cover" && <CoverSlide siteName={siteName} reportMonth={reportMonth} />}
        {current.kind === "glance" && <GlanceSlide summary={deck.summary} />}
        {current.kind === "section" && <SectionSlide page={current.page} />}
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
            title={s.kind === "section" ? s.page.label : s.kind === "cover" ? "Cover" : "At a glance"}
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

function Chrome({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="absolute inset-0 flex flex-col">
      <div style={{ height: "8.6%", background: "var(--blue)" }} />
      <div className="flex min-h-0 flex-1 flex-col px-[4.5%] pt-[2.5%]">
        <p style={{ fontFamily: "var(--serif)", fontSize: "clamp(13px, 2.4vw, 22px)" }}>{title}</p>
        <div style={{ width: "10%", height: 3, background: "var(--sun)", marginTop: "0.8%" }} />
        <div className="mt-[2.2%] min-h-0 flex-1 overflow-hidden">{children}</div>
      </div>
    </div>
  );
}

function CoverSlide({ siteName, reportMonth }: { siteName: string; reportMonth: string }) {
  return (
    <div className="absolute inset-0 flex flex-col">
      <div style={{ height: "8.6%", background: "var(--blue)" }} />
      <div className="flex min-h-0 flex-1">
        <div className="flex w-[40%] flex-col justify-center px-[4%]" style={{ background: "var(--sun)" }}>
          <p
            className="font-semibold"
            style={{ fontSize: "clamp(7px, 1.05vw, 11px)", letterSpacing: "0.18em" }}
          >
            MONTHLY MANAGEMENT REPORT
          </p>
          <div className="mt-[8%] px-[6%] py-[6%]" style={{ background: "var(--blue)", marginLeft: "-6%", marginRight: "-6%" }}>
            <p style={{ fontFamily: "var(--serif)", fontSize: "clamp(14px, 2.6vw, 26px)", color: "#fff" }}>
              {siteName}
            </p>
          </div>
          <p className="mt-[6%]" style={{ fontSize: "clamp(8px, 1.2vw, 13px)" }}>
            {reportMonth}
          </p>
        </div>
        <div className="flex flex-1 items-center justify-center" style={{ background: "var(--wash)" }}>
          <p className="px-6 text-center text-[12px] leading-relaxed" style={{ color: "var(--muted)" }}>
            The site photograph appears here in the .pptx
          </p>
        </div>
      </div>
    </div>
  );
}

function GlanceSlide({ summary }: { summary: CuratedDeck["summary"] }) {
  return (
    <Chrome title="This Month at a Glance">
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

function SectionSlide({ page }: { page: { label: string; rows: Row[] } }) {
  return (
    <Chrome title={page.label}>
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
          <div key={n} className="flex-1 px-[2%] py-[1.5%]" style={{ background: "var(--wash)", borderLeft: "3px solid var(--sun)" }}>
            <p style={{ fontFamily: "var(--serif)", color: "var(--blue)", fontSize: "clamp(10px, 1.7vw, 17px)" }}>
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
              <th key={n} style={{ ...cell, background: "var(--sun)", color: "#fff", textAlign: "left", fontWeight: 600 }}>
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
        {block.imageIds.slice(0, 6).map((id) => (
          <div
            key={id}
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
      <p className="italic" style={{ fontSize: "clamp(5.5px, 0.82vw, 9px)", color: "var(--muted)" }}>
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
// has to show the same figures, in the same shape and the same brand
// colours, as the chart the .pptx will contain. It mirrors the renderer's
// decisions deliberately - uniform colour for a single series, a zero
// baseline on bars, horizontal layout for long category labels - so that
// what the submitter approves here is what gets filed.
// ---------------------------------------------------------------------

const CHART_BASES = ["#264170", "#F7A328"];

function tintHex(hex: string, amount: number): string {
  const n = parseInt(hex.slice(1), 16);
  const mix = (c: number) => Math.round(c + (255 - c) * amount);
  return (
    "#" +
    [mix((n >> 16) & 255), mix((n >> 8) & 255), mix(n & 255)]
      .map((v) => v.toString(16).padStart(2, "0"))
      .join("")
  );
}

/** Mirrors chartPalette() in lib/generate-deck.ts. */
function chartPalette(count: number): string[] {
  return Array.from({ length: Math.max(1, count) }, (_, i) =>
    tintHex(CHART_BASES[i % CHART_BASES.length], Math.min(0.62, Math.floor(i / CHART_BASES.length) * 0.21))
  );
}

const fmt = (n: number) => n.toLocaleString("en-IN");

const trim = (s: string, max: number) => (s.length > max ? s.slice(0, max - 1) + "…" : s);

function ChartView({ block }: { block: Extract<Block, { type: "chart" }> }) {
  const { chartType, categories, series } = block;
  const legend =
    chartType === "pie"
      ? categories.map((c, i) => ({ label: c, color: chartPalette(categories.length)[i] }))
      : series.length > 1
        ? series.map((s, i) => ({ label: s.name, color: chartPalette(series.length)[i] }))
        : [];

  // Grows into the space the other blocks leave, but never below roughly the
  // share the renderer guarantees it (1.5in of a 5.07in content area). Without
  // a floor the flex column squeezed a chart sitting above a ten-row table
  // down to a few unreadable pixels.
  return (
    <div className="flex min-h-0 flex-1 flex-col" style={{ minHeight: "30%" }}>
      {block.title && (
        <p className="mb-[1%]" style={{ fontFamily: "var(--serif)", fontSize: "clamp(6.5px, 1vw, 11px)" }}>
          {block.title}
        </p>
      )}
      <div className="min-h-0 flex-1">
        {chartType === "pie" ? (
          <PieChart categories={categories} values={series[0]?.values ?? []} />
        ) : chartType === "line" ? (
          <LineChart categories={categories} series={series} />
        ) : (
          <BarChart categories={categories} series={series} />
        )}
      </div>
      {legend.length > 0 && (
        <div className="mt-[1%] flex flex-wrap justify-center gap-x-[3%] gap-y-[0.5%]">
          {legend.map((l) => (
            <span
              key={l.label}
              className="flex items-center gap-1"
              style={{ fontSize: "clamp(5px, 0.75vw, 8.5px)", color: "var(--muted)" }}
            >
              <span style={{ width: 6, height: 6, background: l.color, flexShrink: 0 }} />
              {l.label}
            </span>
          ))}
        </div>
      )}
      {block.unit && (
        <p className="mt-[0.5%] text-right" style={{ fontSize: "clamp(4.5px, 0.7vw, 8px)", color: "var(--muted)" }}>
          {block.unit}
        </p>
      )}
    </div>
  );
}

function BarChart({
  categories,
  series,
}: {
  categories: string[];
  series: { name: string; values: number[] }[];
}) {
  const longest = Math.max(...categories.map((c) => c.length));
  const horizontal = series.length === 1 && (longest > 14 || categories.length > 8);
  // A single series is one colour; more than one gets a colour each. The
  // .pptx renderer makes the same choice, for a harder reason - see the
  // chartPalette comment in lib/generate-deck.ts.
  const colors = series.length === 1 ? [CHART_BASES[0]] : chartPalette(series.length);

  // Bars encode value as length, so the scale starts at zero.
  const max = Math.max(0, ...series.flatMap((s) => s.values)) || 1;

  if (horizontal) {
    const rowH = 100 / categories.length;
    return (
      <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
        {categories.map((c, i) => {
          const v = series[0].values[i];
          const w = (v / max) * 205;
          return (
            <g key={c}>
              <text x={92} y={i * rowH + rowH / 2} textAnchor="end" dominantBaseline="middle" fontSize={5} fill="#697784">
                {trim(c, 22)}
              </text>
              <rect x={96} y={i * rowH + rowH * 0.22} width={Math.max(0.5, w)} height={rowH * 0.56} fill={colors[0]} />
              <text x={96 + w + 3} y={i * rowH + rowH / 2} dominantBaseline="middle" fontSize={5} fill="#2D2D2D">
                {fmt(v)}
              </text>
            </g>
          );
        })}
      </svg>
    );
  }

  const colW = 300 / categories.length;
  return (
    <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
      <line x1={18} y1={88} x2={318} y2={88} stroke="#E5E5E4" strokeWidth={0.5} />
      {categories.map((c, i) => {
        const groupX = 18 + i * colW;
        const barW = (colW * 0.62) / series.length;
        const offset = (colW - barW * series.length) / 2;
        return (
          <g key={c}>
            {series.map((s, si) => {
              const h = (s.values[i] / max) * 72;
              const x = groupX + offset + si * barW;
              return (
                <g key={s.name}>
                  <rect x={x} y={88 - h} width={Math.max(0.5, barW * 0.88)} height={Math.max(0.3, h)} fill={colors[si]} />
                  <text x={x + (barW * 0.88) / 2} y={88 - h - 2} textAnchor="middle" fontSize={4.5} fill="#2D2D2D">
                    {fmt(s.values[i])}
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
  const colors = chartPalette(series.length);
  const all = series.flatMap((s) => s.values);
  // A line encodes value by position, so unlike bars it need not start at
  // zero - the renderer leaves this axis automatic for the same reason.
  const lo = Math.min(...all);
  const hi = Math.max(...all);
  const pad = (hi - lo) * 0.15 || Math.abs(hi) * 0.1 || 1;
  const min = lo - pad;
  const span = hi + pad - min || 1;
  const x = (i: number) => 22 + (i * 292) / Math.max(1, categories.length - 1);
  const y = (v: number) => 82 - ((v - min) / span) * 72;

  return (
    <svg viewBox="0 0 320 100" preserveAspectRatio="xMidYMid meet" className="h-full w-full">
      <line x1={18} y1={82} x2={318} y2={82} stroke="#E5E5E4" strokeWidth={0.5} />
      {series.map((s, si) => (
        <g key={s.name}>
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
        <text key={c} x={x(i)} y={92} textAnchor="middle" fontSize={5} fill="#697784">
          {trim(c, 10)}
        </text>
      ))}
    </svg>
  );
}

function PieChart({ categories, values }: { categories: string[]; values: number[] }) {
  const colors = chartPalette(categories.length);
  const total = values.reduce((a, b) => a + b, 0) || 1;
  const cx = 50;
  const cy = 50;

  // Each slice's start angle, accumulated up front rather than carried in a
  // variable across the map below - reassigning one during render is exactly
  // what react-hooks/immutability forbids.
  const starts: number[] = [];
  values.reduce((acc, v) => {
    starts.push(acc);
    return acc + (v / total) * Math.PI * 2;
  }, -Math.PI / 2);

  // A filled pie with labels outside, matching what the .pptx renders and
  // for the same reason: no single label colour stays readable against both
  // the dark and the pale ends of the palette, so the labels sit on white.
  const arc = (a0: number, a1: number) => {
    const pt = (r: number, a: number) => `${cx + r * Math.cos(a)} ${cy + r * Math.sin(a)}`;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    return `M${cx} ${cy} L${pt(31, a0)} A31 31 0 ${large} 1 ${pt(31, a1)} Z`;
  };

  return (
    <svg viewBox="0 0 100 100" className="h-full w-full">
      {values.map((v, i) => {
        const a0 = starts[i];
        const a1 = a0 + (v / total) * Math.PI * 2;
        const mid = (a0 + a1) / 2;
        return (
          <g key={categories[i]}>
            <path d={arc(a0, a1)} fill={colors[i]} />
            {v / total > 0.03 && (
              <text
                x={cx + 36 * Math.cos(mid)}
                y={cy + 36 * Math.sin(mid)}
                textAnchor="middle"
                dominantBaseline="middle"
                fontSize={4.5}
                fill="#2D2D2D"
              >
                {fmt(v)}
              </text>
            )}
          </g>
        );
      })}
    </svg>
  );
}
