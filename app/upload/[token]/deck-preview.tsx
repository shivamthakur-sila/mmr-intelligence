"use client";

import { useState } from "react";
import type { CuratedDeck, Block } from "@/lib/curate";
import {
  CONTENT_W,
  KPI_CARD_MAX_W,
  KPI_ROW_H,
  KPI_STRIP_H,
  KPI_STACK_CARD_H,
  KPI_STACK_GAP,
  SPLIT_GUTTER,
  TEXT_PANEL_MAX_W,
  TEXT_PANEL_PAD,
  deckOutline,
  identifierColumns,
  numericColumns,
  placeRows,
  splitRightW,
  tableGeometry,
  tableType,
  type ContentsChapter,
  type OutlineSlide,
  type PhotoGeometry,
  type Row,
} from "@/lib/deck-layout";
import { chartPalette, decimalsNeeded, formatNumber, indianFigures, niceAxis } from "@/lib/chart-style";

/** Left edge of the content area, in the renderer's inches (CONTENT_X there). */
const CONTENT_X = 0.6;

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
  photos = {},
  initialSlide = 0,
}: {
  deck: CuratedDeck;
  siteName: string;
  reportMonth: string;
  /** The site photo the deck's cover uses, or null for SILA's building. */
  coverImage?: string | null;
  /** Small copies of the photos the deck's grids use, by id. */
  photos?: Record<string, string>;
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
        // Sizes inside are in container units of this box, so type and spacing
        // keep the deck's proportions at any width.
        style={{
          aspectRatio: "16 / 9",
          background: "#fff",
          border: "1px solid var(--hairline)",
          containerType: "inline-size",
          fontFamily: '"Trebuchet MS", var(--sans)',
        }}
      >
        {current.kind === "cover" && <CoverSlide siteName={siteName} reportMonth={reportMonth} photo={coverImage} />}
        {current.kind === "glance" && <GlanceSlide summary={deck.summary} page={i + 1} />}
        {current.kind === "contents" && <ContentsSlide chapters={current.chapters} reportMonth={reportMonth} page={i + 1} />}
        {current.kind === "divider" && <DividerSlide number={current.number} chapter={current.chapter} page={i + 1} photo={coverImage} />}
        {current.kind === "section" && <SectionSlide page={current} pageNumber={i + 1} photos={photos} />}
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
      style={{ ...box(12.586, 6.753, 0.747, 0.747), background: "var(--sun)", color: CHARCOAL, fontSize: `${(12 * 7.5) / 72}cqw` }}
    >
      {String(n).padStart(2, "0")}
    </div>
  );
}

function Chrome({
  title,
  page,
  children,
  bare = false,
}: {
  title: string;
  page: number;
  children: React.ReactNode;
  /** True when the children place themselves in slide inches. */
  bare?: boolean;
}) {
  return (
    <div className="absolute inset-0">
      <Band />
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/sila-fm-logo-color.png" alt="SILA" style={{ ...box(11.683, 0.9, 1.4, 0.595), objectFit: "contain" }} />
      <p
        className="flex items-center overflow-hidden whitespace-nowrap"
        style={{ ...box(0.7, 0.92, 10.6, 0.66), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: `${(32 * 7.5) / 72}cqw` }}
      >
        {title}
      </p>
      <div style={{ ...box(0.7, 1.7, 1.11, 0.04), background: "var(--sun)" }} />
      <div style={{ ...box(13.226, 3.667, 0.107, 0.667), background: "#FFAF00" }} />
      {bare ? (
        children
      ) : (
        <div className="flex flex-col overflow-hidden" style={box(0.6, 1.95, 12.13, 4.75)}>
          {children}
        </div>
      )}
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

/** Mirrors renderGlance in generate-deck.ts: headline, figure tiles, facts in two columns. */
function GlanceSlide({ summary, page }: { summary: CuratedDeck["summary"]; page: number }) {
  const headLines = Math.max(1, Math.ceil(summary.headline.length / 105));
  const kpis = summary.kpis ?? [];
  const points = summary.points.slice(0, 6);
  const cols = points.length >= 3 ? 2 : 1;
  const perCol = Math.ceil(points.length / cols);
  let y = 1.95 + 0.22 + headLines * 0.3 + (kpis.length > 0 ? KPI_ROW_H + 0.3 : 0);
  const ruleY = y;
  y += 0.22;
  const rowH = Math.min(0.85, (6.7 - y) / Math.max(1, perCol));
  const colW = (CONTENT_W - 0.5 * (cols - 1)) / cols;
  return (
    <Chrome title="This Month at a Glance" page={page} bare>
      <p style={{ ...box(CONTENT_X, 1.95, CONTENT_W, 0.12 + headLines * 0.3), fontFamily: "var(--serif)", color: "var(--blue)", fontSize: pt(17), lineHeight: 1.2 }}>
        {indianFigures(summary.headline)}
      </p>
      {kpis.length > 0 && (
        <div style={box(CONTENT_X, 1.95 + 0.22 + headLines * 0.3, CONTENT_W, KPI_ROW_H)}>
          <BlockView block={{ type: "kpis", items: kpis }} w={CONTENT_W} photos={{}} />
        </div>
      )}
      <div style={{ ...box(CONTENT_X, ruleY, CONTENT_W, 0.015), background: "var(--hairline)" }} />
      {points.map((p, n) => {
        const col = Math.floor(n / perCol);
        const x = CONTENT_X + col * (colW + 0.5);
        const py = y + (n % perCol) * rowH;
        return (
          <div key={n}>
            <div style={{ ...box(x, py + 0.06, 0.07, Math.min(0.42, rowH - 0.16)), background: "var(--sun)" }} />
            <p style={{ ...box(x + 0.22, py, colW - 0.22, rowH - 0.08), fontSize: pt(12), color: SLATE, lineHeight: 1.25, overflow: "hidden" }}>{indianFigures(p)}</p>
          </div>
        );
      })}
    </Chrome>
  );
}

function slideTitle(s: OutlineSlide): string {
  switch (s.kind) {
    case "cover": return "Cover";
    case "glance": return "At a glance";
    case "contents": return "Contents";
    case "divider": return `${String(s.number).padStart(2, "0")} ${s.chapter}`;
    case "closing": return "Thank you";
    case "section": return s.label;
  }
}

/** Mirrors renderContents: numbered chapters, their sections, the page each starts on. */
function ContentsSlide({ chapters, reportMonth, page }: { chapters: ContentsChapter[]; reportMonth: string; page: number }) {
  const cols = chapters.length > 5 ? 2 : 1;
  const colW = (CONTENT_W - 0.6 * (cols - 1)) / cols;
  const perCol = Math.ceil(chapters.length / cols);
  const rowH = Math.min(0.95, 4.75 / Math.max(1, perCol));
  const sq = Math.min(0.5, rowH - 0.25);
  return (
    <Chrome title={`Contents — ${prettyMonth(reportMonth)}`} page={page} bare>
      {chapters.map((ch, i) => {
        const x = CONTENT_X + Math.floor(i / perCol) * (colW + 0.6);
        const y = 1.95 + (i % perCol) * rowH;
        return (
          <div key={i}>
            <div
              className="flex items-center justify-center"
              style={{ ...box(x, y + 0.05, sq, sq), background: "var(--sun)", fontFamily: "var(--serif)", color: CHARCOAL, fontSize: pt(14) }}
            >
              {String(ch.number).padStart(2, "0")}
            </div>
            <p className="flex items-center whitespace-nowrap" style={{ ...box(x + sq + 0.2, y, colW - sq - 1.0, 0.36), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: pt(15) }}>
              {ch.name}
            </p>
            <p style={{ ...box(x + sq + 0.2, y + 0.36, colW - sq - 1.0, rowH - 0.5), color: MUTED, fontSize: pt(10), overflow: "hidden" }}>
              {ch.sections.join("  ·  ")}
            </p>
            <p className="flex items-center justify-end" style={{ ...box(x + colW - 0.7, y, 0.7, 0.36), fontFamily: "var(--serif)", color: "var(--blue)", fontSize: pt(14) }}>
              {String(ch.page).padStart(2, "0")}
            </p>
            <div style={{ ...box(x, y + rowH - 0.1, colW, 0.01), background: "var(--hairline)" }} />
          </div>
        );
      })}
    </Chrome>
  );
}

/** Mirrors renderDivider: the deck's photo, an orange block with the chapter, a navy strip. */
function DividerSlide({ number, chapter, page, photo }: { number: number; chapter: string; page: number; photo: string | null }) {
  return (
    <div className="absolute inset-0">
      <Band />
      <PanelPhoto photo={photo} />
      {photo && <LogoPlate x={10.4} y={0.75} w={2.3} h={1.15} />}
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src="/sila-fm-logo-color.png" alt="SILA" style={{ ...box(10.55, 0.9, 2.0, 0.85), objectFit: "contain" }} />
      <div style={{ ...box(0, 4.64, 9.28, 2.13), background: "var(--sun)" }} />
      <div style={{ ...box(0, 6.77, 9.29, 0.73), background: "var(--blue)" }} />
      <p className="flex items-center" style={{ ...box(0.76, 4.78, 2, 0.45), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: pt(18), letterSpacing: "0.15em" }}>
        {String(number).padStart(2, "0")}
      </p>
      <p className="flex items-center" style={{ ...box(0.76, 5.2, 8.2, 1.4), fontFamily: "var(--serif)", color: CHARCOAL, fontSize: pt(36), letterSpacing: "0.12em", lineHeight: 1.1 }}>
        {chapter.toUpperCase()}
      </p>
      <PageSquare n={page} />
    </div>
  );
}

/** A font size in points on the 13.333in slide, as a share of the slide's width. */
const pt = (n: number) => `${(n * 7.5) / 72}cqw`;
/** A length in inches on the slide, as a share of its width. */
const inch = (n: number) => `${n * 7.5}cqw`;
const SLATE = "#2D2D2D";
const MUTED = "#697784";

// The same status fills the renderer uses: the source's word stays, the
// colour only marks it.
const STATUS_FILLS: { test: RegExp; fill: string; color: string }[] = [
  { test: /^(completed?|done|closed|ok|okay|working|valid|available|resolved|approved|active|renewed)$/i, fill: "#E4F2E7", color: "#1E6B34" },
  { test: /^(pending|open|in progress|wip|ongoing|scheduled|under process|due)$/i, fill: "#FDF0D9", color: "#8A5A00" },
  { test: /^(overdue|expired|not done|not working|failed|rejected|breakdown|out of order)$/i, fill: "#F9E1E1", color: "#9B2C2C" },
];

/** A section slide: every block placed where the renderer places it. */
function SectionSlide({
  page,
  pageNumber,
  photos,
}: {
  page: { label: string; rows: Row[] };
  pageNumber: number;
  photos: Record<string, string>;
}) {
  return (
    <Chrome title={page.label} page={pageNumber} bare>
      {placeRows(page.rows).map(({ row, y, h, geo }, n) =>
        row.kind === "full" ? (
          <div key={n} style={box(CONTENT_X, y, CONTENT_W, h)}>
            <BlockView block={row.block} w={CONTENT_W} geo={geo} photos={photos} />
          </div>
        ) : (
          <div key={n}>
            <div style={box(CONTENT_X, y, row.leftW, h)}>
              <BlockView block={row.left} w={row.leftW} photos={photos} />
            </div>
            <div style={box(CONTENT_X + row.leftW + SPLIT_GUTTER, y, splitRightW(row), h)}>
              <BlockView block={row.right} w={splitRightW(row)} geo={row.right.type === "photos" ? geo : undefined} photos={photos} />
            </div>
          </div>
        )
      )}
    </Chrome>
  );
}

function BlockView({
  block,
  w,
  geo,
  photos,
}: {
  block: Block;
  w: number;
  geo?: PhotoGeometry;
  photos: Record<string, string>;
}) {
  if (block.type === "kpis") {
    const shown = block.items.slice(0, 4);
    const stacked = w < CONTENT_W / 2;
    const cardW = stacked ? w : Math.min(KPI_CARD_MAX_W, (w - 0.22 * (shown.length - 1)) / shown.length);
    const cardH = stacked ? KPI_STACK_CARD_H : block.compact ? KPI_STRIP_H : KPI_ROW_H;
    return (
      <div className={stacked ? "flex flex-col" : "flex"} style={{ gap: inch(stacked ? KPI_STACK_GAP : 0.22) }}>
        {shown.map((k, n) => (
          <div
            key={n}
            className="flex flex-col justify-center overflow-hidden"
            style={{
              width: inch(cardW),
              height: inch(cardH),
              background: "#fff",
              border: "1px solid var(--hairline)",
              borderLeft: `${inch(0.07)} solid var(--sun)`,
              paddingLeft: inch(0.17),
            }}
          >
            <p className="whitespace-nowrap" style={{ fontFamily: "var(--serif)", color: "var(--blue)", fontSize: pt(block.compact ? 20 : stacked ? 26 : 30), lineHeight: 1.1 }}>
              {indianFigures(k.value)}
            </p>
            <p style={{ fontSize: pt(8.5), color: MUTED, letterSpacing: "0.1em", marginTop: inch(0.06) }}>{k.label.toUpperCase()}</p>
          </div>
        ))}
      </div>
    );
  }

  if (block.type === "table") {
    const tg = tableGeometry(block, w);
    const type = tableType(block);
    const numeric = numericColumns(block);
    const ids = identifierColumns(block);
    const cell = (c: number): React.CSSProperties => ({
      fontSize: pt(type.body),
      padding: `${inch(type.pad * 0.3)} ${inch(0.1)}`,
      lineHeight: 1.15,
      borderBottom: "1px solid var(--hairline)",
      textAlign: numeric[c] ? "right" : "left",
      color: SLATE,
      whiteSpace: "pre-line",
      verticalAlign: "middle",
    });
    const bands = new Map((block.groups ?? []).map((g) => [Math.min(g.afterRow, block.rows.length), g.label]));
    const out: React.ReactNode[] = [];
    const band = (ri: number) => {
      if (!bands.has(ri)) return;
      out.push(
        <tr key={`band-${ri}`}>
          <td colSpan={block.headers.length} style={{ ...cell(0), textAlign: "left", background: "#E8EDF4", color: "var(--blue)", fontWeight: 700 }}>
            {bands.get(ri)}
          </td>
        </tr>
      );
    };
    block.rows.forEach((row, ri) => {
      band(ri);
      out.push(
        <tr key={`row-${ri}`}>
          {row.map((c, ci) => {
            const status = STATUS_FILLS.find((s) => s.test.test(String(c).trim()));
            return (
              <td key={ci} style={{ ...cell(ci), ...(status ? { background: status.fill, color: status.color, fontWeight: 700 } : {}) }}>
                {numeric[ci] && !ids[ci] ? indianFigures(String(c)) : c}
              </td>
            );
          })}
        </tr>
      );
    });
    band(block.rows.length);
    return (
      <div>
        <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed" }}>
          <colgroup>
            {tg.widths.map((cw, n) => (
              <col key={n} style={{ width: `${(cw / w) * 100}%` }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              {block.headers.map((hd, n) => (
                <th key={n} style={{ ...cell(n), fontSize: pt(type.header), background: "var(--blue)", color: "#fff", fontWeight: 700 }}>
                  {hd}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>{out}</tbody>
        </table>
        {block.footnote && (
          <p style={{ fontSize: pt(9), fontStyle: "italic", color: MUTED, marginTop: inch(0.04) }}>{block.footnote}</p>
        )}
      </div>
    );
  }

  if (block.type === "chart") {
    return (
      <div className="flex h-full flex-col">
        <div className="flex min-h-0 flex-1 flex-col">
          <ChartView block={block} />
        </div>
        {block.footnote && <p style={{ fontSize: pt(9), fontStyle: "italic", color: MUTED }}>{block.footnote}</p>}
      </div>
    );
  }

  if (block.type === "photos") {
    if (!geo) return null;
    const ids = block.imageIds.slice(0, 6);
    return (
      <div className="flex h-full justify-center">
        <div className="grid content-start" style={{ gridTemplateColumns: `repeat(${geo.cols}, ${inch(geo.w)})`, gap: inch(0.18) }}>
          {ids.map((id) =>
            photos[id] ? (
              // eslint-disable-next-line @next/next/no-img-element
              <img key={id} src={photos[id]} alt="" style={{ width: inch(geo.w), height: inch(geo.h), objectFit: "cover", border: `1px solid ${CHARCOAL}` }} />
            ) : (
              <div key={id} style={{ width: inch(geo.w), height: inch(geo.h), background: "var(--wash)" }} />
            )
          )}
        </div>
      </div>
    );
  }

  // Beside photos or a chart, text sits in a pale panel with an orange edge.
  const panel: React.CSSProperties =
    w < TEXT_PANEL_MAX_W
      ? { height: "100%", background: "#F3F5F8", borderLeft: `${inch(0.06)} solid var(--sun)`, padding: inch(TEXT_PANEL_PAD) }
      : {};

  if (block.type === "bullets") {
    return (
      <div style={panel}>
        <ul
          style={{
            paddingLeft: inch(0.3), fontSize: pt(12), color: SLATE, lineHeight: 1.25, listStyle: "disc",
            ...(block.columns === 2 ? { columnCount: 2, columnGap: inch(0.3) } : {}),
          }}
        >
          {block.items.map((t, n) => (
            <li key={n} style={{ marginBottom: w < TEXT_PANEL_MAX_W || block.columns === 2 ? pt(4) : 0, breakInside: "avoid" }}>{t}</li>
          ))}
        </ul>
      </div>
    );
  }

  if (block.type === "note") {
    return (
      <p className="flex h-full items-center" style={{ fontSize: pt(10), color: SLATE, border: "1px dashed var(--sun)", padding: `0 ${inch(0.16)}` }}>
        {block.text}
      </p>
    );
  }

  return (
    <div style={panel}>
      <p style={{ fontSize: pt(12.5), color: SLATE, lineHeight: 1.3 }}>{block.text}</p>
    </div>
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
