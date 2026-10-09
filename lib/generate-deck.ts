import pptxgen from "pptxgenjs";
import sharp from "sharp";
import fs from "fs";
import os from "os";
import path from "path";
import type { CuratedDeck, Block } from "./curate";
import { axisDecimals, chartPalette, decimalsNeeded, indianFigures, niceAxis, numberFormatCode } from "./chart-style";
import {
  CONTENT_TOP,
  CONTENT_W,
  FOOTNOTE_SIZE,
  KPI_CARD_MAX_W,
  KPI_ROW_H,
  KPI_STACK_CARD_H,
  KPI_STACK_GAP,
  SPLIT_GUTTER,
  type PhotoGeometry,
  type Row,
  blockHeight,
  footnoteHeight,
  identifierColumns,
  numericColumns,
  placeRows,
  sectionSlides,
  splitRightW,
  tableGeometry,
  tableType,
} from "./deck-layout";

// FM service-line palette. The parent SILA guide assigns Blue #264170 to
// Facility Management and pairs service colours with Sunshine rather than
// Charcoal, so this is the correct palette for an FM deliverable.
const FM_BLUE = "264170";
const SUNSHINE = "F7A328";
const SLATE = "2D2D2D";
// The reference deck sets titles, cover type and the page number in charcoal.
const CHARCOAL = "3C3C3B";
// The small tab on the right edge of every content slide in the reference.
const EDGE_TAB = "FFAF00";
const MUTED = "697784";
const HAIRLINE = "E5E5E4";
const WHITE = "FFFFFF";

// Brand fonts are Merriweather (headings) / Avenir Next (body). Neither is
// installed on a client's machine, pptxgenjs cannot embed fonts, and a missing
// font renders as an ugly substitute. The reference deck (Sila x Sattva) uses
// Georgia for titles and Trebuchet MS for body, both of which ship with every
// copy of Windows, Office and macOS, so the deck reads the same everywhere.
const HEAD_FONT = "Georgia";
const BODY_FONT = "Trebuchet MS";

const SLIDE_W = 13.333;
const CONTENT_X = 0.6;

type ImageRecord = {
  id: string;
  file: string;
  slideNumber: number;
  slideTitle: string;
  width?: number;
  height?: number;
};

// The cover, divider and closing photo panel of the reference deck: inset
// under the 0.75in band, with white margins left, right and below.
const PHOTO_PANEL = { x: 0.76, y: 0.75, w: 11.83, h: 6.0 };

/** SILA's own black-and-white building, from the official PPT template. */
function fallbackBuilding(): string {
  const p = path.join(process.cwd(), "lib/assets", "bw-building.jpg");
  if (!fs.existsSync(p)) {
    throw new Error(`Fallback cover image missing: ${p}. lib/assets must ship with the deployment.`);
  }
  return p;
}

// Below about 110 pixels per placed inch a photo prints visibly soft. The
// cover photo is drawn 11.83in wide, so it needs about 1,300 pixels across
// after cropping to the panel; a 744px cover was blown up to fill it.
const MIN_PANEL_PX = 1300;

/** Pixels across a photo keeps once centre-cropped to the panel's ratio. */
function croppedWidth(im: ImageRecord): number {
  if (!im.width || !im.height) return 0;
  const ratio = PHOTO_PANEL.w / PHOTO_PANEL.h;
  return Math.min(im.width, im.height * ratio);
}

/**
 * The site photograph for the cover and closing slides, or undefined when
 * SILA's building stands in.
 *
 * Only the model's nomination is used: it chose knowing which slide each photo
 * came from, so it can tell a building exterior from a plumber at work. When
 * the nominated photo is too soft for the panel, the fallback is SILA's own
 * building, never the "sharpest" other photo - on a real deck that rule put a
 * pipe repair on the cover.
 */
/**
 * A small black-and-white copy of the photo the cover and closing slides use,
 * for the browser preview. Site photos are in private storage the browser
 * cannot read, so without this the preview showed SILA's building even when
 * the deck had the site's own photo - and the submitter approves the preview.
 * Null when SILA's building is used; the preview has that one itself.
 */
export async function coverPreviewDataUrl(
  deck: CuratedDeck,
  images: ImageRecord[],
  sessionDir: string
): Promise<string | null> {
  const photo = chooseCoverPhoto(deck, images, sessionDir);
  if (!photo) return null;
  try {
    const jpeg = await sharp(path.join(sessionDir, photo.file))
      .resize({ width: 640, height: Math.round(640 / (PHOTO_PANEL.w / PHOTO_PANEL.h)), fit: "cover", position: "centre" })
      .grayscale()
      .jpeg({ quality: 70 })
      .toBuffer();
    return `data:image/jpeg;base64,${jpeg.toString("base64")}`;
  } catch {
    return null;
  }
}

/**
 * Small copies of every photo the deck's grids use, keyed by id, for the
 * browser preview. It drew grey "photo" boxes, so the submitter approved
 * photo slides without seeing which photographs were on them.
 */
export async function photoPreviewDataUrls(
  deck: CuratedDeck,
  images: ImageRecord[],
  sessionDir: string
): Promise<Record<string, string>> {
  const used = new Set(
    deck.sections.flatMap((s) => s.blocks.flatMap((b) => (b.type === "photos" ? b.imageIds : [])))
  );
  const out: Record<string, string> = {};
  for (const im of images) {
    if (!used.has(im.id)) continue;
    try {
      const jpeg = await sharp(path.join(sessionDir, im.file))
        .resize({ width: 360, height: 360, fit: "inside", withoutEnlargement: true })
        .jpeg({ quality: 68 })
        .toBuffer();
      out[im.id] = `data:image/jpeg;base64,${jpeg.toString("base64")}`;
    } catch {
      // A preview without one thumbnail still shows the tile.
    }
  }
  return out;
}

export function chooseCoverPhoto(
  deck: CuratedDeck,
  images: ImageRecord[],
  sessionDir: string
): ImageRecord | undefined {
  const usable = images.filter(
    (im) => fs.existsSync(path.join(sessionDir, im.file)) && croppedWidth(im) >= MIN_PANEL_PX
  );
  return usable.find((im) => im.id === deck.coverImageId);
}


/** Turns "2026-08" into "August 2026". Anything else passes through. */
function prettyMonth(value: string): string {
  const m = /^(\d{4})-(\d{2})$/.exec(value.trim());
  if (!m) return value;
  const names = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"];
  const idx = parseInt(m[2], 10) - 1;
  return names[idx] ? `${names[idx]} ${m[1]}` : value;
}

/**
 * Path to a brand logo. Throws rather than returning null: the caller used to
 * skip the logo when the file could not be found, so a deck would go to a
 * client with no branding and nothing anywhere to say why. Preflight caught
 * exactly that happening on fourteen slides.
 */
function logo(white: boolean): string {
  const name = white ? "sila-fm-logo-white.png" : "sila-fm-logo-color.png";
  const p = path.join(process.cwd(), "lib/assets", name);
  if (!fs.existsSync(p)) {
    throw new Error(
      `Brand logo missing: ${p}. Deck generation runs from ${process.cwd()}; ` +
        `lib/assets must ship with the deployment.`
    );
  }
  return p;
}

/**
 * Downscaled copies of both logos, prepared once per build.
 *
 * pptxgenjs embeds an image per addImage call rather than once per file, so the
 * full-size 180KB logo was stored 18 times - 3.3MB of one picture, more than
 * half the deck. Neither is ever drawn more than 3in wide.
 */
const preparedLogos: { white: string | null; color: string | null } = { white: null, color: null };
async function prepareLogo(): Promise<void> {
  for (const white of [true, false]) {
    const key = white ? "white" : "color";
    const out = path.join(os.tmpdir(), `sila-fm-logo-${key}-480.png`);
    try {
      if (!fs.existsSync(out)) await sharp(logo(white)).resize({ width: 480 }).png().toFile(out);
      preparedLogos[key] = out;
    } catch {
      preparedLogos[key] = null; // fall back to the full-size original
    }
  }
}
function brandLogo(white = true): string {
  return preparedLogos[white ? "white" : "color"] ?? logo(white);
}

// Logo aspect ratios (width / height) of the files in lib/assets, so a logo is
// placed at its own proportions and never stretched.
const COLOR_LOGO_RATIO = 1545 / 657;

// The page furniture of the reference deck, scaled from its 20in page to
// pptxgenjs's 13.333in WIDE layout (factor 0.667) and measured off the PDF:
// a 0.75in navy band, an orange page square flush in the bottom-right corner
// with the number in charcoal, the colour logo on white under the band, and a
// small orange tab on the right edge.
const BAND_H = 0.75;
const PAGE_SQ = 0.747;
const TITLE_RULE_Y = 1.7;

/** The page-number square, shared by every slide that has one. */
function pageSquare(slide: pptxgen.Slide, pageNum: number) {
  const x = SLIDE_W - PAGE_SQ;
  const y = 7.5 - PAGE_SQ;
  slide.addShape("rect", { x, y, w: PAGE_SQ, h: PAGE_SQ, fill: { color: SUNSHINE }, line: { type: "none" } });
  slide.addText(String(pageNum).padStart(2, "0"), {
    x, y, w: PAGE_SQ, h: PAGE_SQ, align: "center", valign: "middle", margin: 0,
    fontFace: BODY_FONT, fontSize: 12, bold: true, color: CHARCOAL,
  });
}

function band(slide: pptxgen.Slide) {
  slide.addShape("rect", { x: 0, y: 0, w: SLIDE_W, h: BAND_H, fill: { color: FM_BLUE }, line: { type: "none" } });
}

function chrome(slide: pptxgen.Slide, title: string, pageNum: number) {
  slide.background = { color: WHITE };
  band(slide);

  // The colour logo sits on white under the band, top-right. The white logo
  // inside the band read as a small mark; this is the reference's placement.
  const logoW = 1.4;
  slide.addImage({ path: brandLogo(false), x: SLIDE_W - 0.25 - logoW, y: 0.9, w: logoW, h: logoW / COLOR_LOGO_RATIO });

  // Georgia at about three times the body size, as the reference sets it, in
  // charcoal; the fixed-length orange rule under it does not follow the
  // title's length.
  slide.addText(title, {
    x: CONTENT_X, y: 0.92, w: SLIDE_W - CONTENT_X - logoW - 0.6, h: 0.66,
    fontFace: HEAD_FONT, fontSize: 32, color: CHARCOAL, valign: "middle", fit: "shrink",
  });
  slide.addShape("rect", { x: CONTENT_X + 0.1, y: TITLE_RULE_Y, w: 1.11, h: 0.04, fill: { color: SUNSHINE }, line: { type: "none" } });

  slide.addShape("rect", { x: SLIDE_W - 0.107, y: 3.667, w: 0.107, h: 0.667, fill: { color: EDGE_TAB }, line: { type: "none" } });
  pageSquare(slide, pageNum);
}

/** One line of the contents list: orange tick bar, then the section name. */
function slide_contentsRow(slide: pptxgen.Slide, label: string, x: number, y: number, w: number) {
  slide.addShape("rect", { x, y: y + 0.08, w: 0.1, h: 0.22, fill: { color: SUNSHINE }, line: { type: "none" } });
  slide.addText(label, {
    x: x + 0.3, y, w: w - 0.3, h: 0.38,
    fontFace: BODY_FONT, fontSize: 13, color: SLATE, valign: "middle",
  });
}

function renderKpis(
  slide: pptxgen.Slide,
  items: { label: string; value: string }[],
  x: number,
  y: number,
  w: number,
  stacked: boolean
) {
  // Carded like the house decks: a bordered panel with an orange edge and
  // the figure set large. In a row each card is capped in width and the row
  // starts at the left edge; beside a table or chart the cards stack into a
  // column. Values are regrouped the Indian way for display only.
  const shown = items.slice(0, 4);
  const card = (kx: number, ky: number, kw: number, kh: number, kpi: { label: string; value: string }) => {
    slide.addShape("rect", { x: kx, y: ky, w: kw, h: kh, fill: { color: WHITE }, line: { color: HAIRLINE, width: 1 } });
    slide.addShape("rect", { x: kx, y: ky, w: 0.07, h: kh, fill: { color: SUNSHINE }, line: { type: "none" } });
    slide.addText(indianFigures(kpi.value), {
      x: kx + 0.24, y: ky + 0.1, w: kw - 0.36, h: kh * 0.55,
      fontFace: HEAD_FONT, fontSize: stacked ? 26 : 30, color: FM_BLUE, valign: "middle", fit: "shrink",
    });
    slide.addText(kpi.label.toUpperCase(), {
      x: kx + 0.26, y: ky + kh * 0.62, w: kw - 0.38, h: kh * 0.3,
      fontFace: BODY_FONT, fontSize: 8.5, color: MUTED, charSpacing: 1.2, valign: "middle", fit: "shrink",
    });
  };
  if (stacked) {
    shown.forEach((kpi, i) => card(x, y + i * (KPI_STACK_CARD_H + KPI_STACK_GAP), w, KPI_STACK_CARD_H, kpi));
    return;
  }
  const gap = 0.22;
  const cw = Math.min(KPI_CARD_MAX_W, (w - gap * (shown.length - 1)) / shown.length);
  shown.forEach((kpi, i) => card(x + i * (cw + gap), y, cw, KPI_ROW_H, kpi));
}

/**
 * Renders a curated chart into the given box.
 *
 * Chart geometry, unlike `addImage`, is honoured exactly as given — checked by
 * drawing a box on the requested rectangle and confirming the plot landed on
 * it — so no cropping dance is needed here.
 */
function renderChart(
  slide: pptxgen.Slide,
  b: Extract<Block, { type: "chart" }>,
  y: number,
  h: number,
  x: number = CONTENT_X,
  w: number = CONTENT_W,
  /** True when the chart shares its band with the table of the same figures. */
  compact = false
) {
  const all = b.series.flatMap((s) => s.values);
  // Shown exactly as the source states them. A fixed "#,##0" rounded 12.5 to
  // 13 on both the axis and the data labels - a figure that is not in the
  // source, which is the one thing this system must never print.
  const valueFormat = numberFormatCode(decimalsNeeded(all));
  const unit = typeof b.unit === "string" ? b.unit.trim() : "";

  // A footnote takes its own lines under the plot.
  const footH = footnoteHeight(b.footnote, w);
  if (b.footnote) {
    slide.addText(b.footnote, {
      x, y: y + h - footH + 0.04, w, h: footH, margin: [0, 2, 0, 2],
      fontFace: BODY_FONT, fontSize: FOOTNOTE_SIZE, italic: true, color: MUTED, valign: "top",
    });
  }

  const common = {
    x,
    y,
    w,
    h: h - footH,
    showTitle: !!b.title,
    title: b.title ?? "",
    titleColor: SLATE,
    titleFontFace: HEAD_FONT,
    titleFontSize: 13,
    catAxisLabelColor: MUTED,
    valAxisLabelColor: MUTED,
    catAxisLabelFontFace: BODY_FONT,
    valAxisLabelFontFace: BODY_FONT,
    // Half-width charts get smaller axis type: at 10pt a category label like
    // "Housekeeping" broke mid-word under its own bar.
    catAxisLabelFontSize: compact ? 8 : 10,
    valAxisLabelFontSize: compact ? 8 : 10,
    dataLabelFormatCode: valueFormat,
    dataLabelFontFace: BODY_FONT,
    legendFontFace: BODY_FONT,
    legendColor: MUTED,
    legendFontSize: 10,
    valGridLine: { color: HAIRLINE, size: 0.5 },
    catGridLine: { style: "none" as const },
  };

  if (b.chartType === "pie") {
    const cats = b.categories;
    // A pie rather than a doughnut, because only a pie honours a label
    // position of "outEnd" — on a doughnut the option is ignored and every
    // label is forced back inside its slice. That matters here: slices
    // further round the palette are pale, so labels inside them have to be
    // dark, and labels inside the first two slices have to be light. No
    // single label colour works. Outside, on white, one colour always does.
    // Both checked by rendering.
    slide.addChart(
      "pie",
      [{ name: b.series[0]?.name ?? "", labels: cats, values: b.series[0]?.values ?? [] }],
      {
        ...common,
        // A pie has no value axis to carry the unit, so it goes on the title.
        // The unit used to reach the preview and never the deck.
        showTitle: !!(b.title || unit),
        title: b.title ? (unit ? `${b.title} (${unit})` : b.title) : unit,
        // One colour per slice — see chartPalette.
        chartColors: chartPalette(cats.length),
        // A white edge between slices, so two neighbouring pale slices
        // still read as two.
        dataBorder: { pt: 1, color: "FFFFFF" },
        showLegend: true,
        legendPos: "r",
        showValue: !compact,
        dataLabelPosition: "outEnd",
        dataLabelColor: SLATE,
        dataLabelFontSize: 10,
      }
    );
    return;
  }

  if (b.chartType === "line") {
    // A line encodes value by position rather than length, so a baseline
    // above zero is legitimate here — and necessary. Left automatic the axis
    // ran 0 to 100,000 against consumption moving 27,240 to 28,920, drawing
    // the month-on-month trend as a dead flat rule. Padding the real range
    // is what makes the shape visible, and it matches the scale the browser
    // preview already uses. Bars are the opposite case: see valAxisMinVal
    // below.
    const axis = niceAxis(all);

    slide.addChart(
      "line",
      b.series.map((s) => ({ name: s.name, labels: b.categories, values: s.values })),
      {
        ...common,
        chartColors: chartPalette(b.series.length),
        lineDataSymbol: "circle",
        lineSize: 2,
        lineSmooth: false,
        showLegend: b.series.length > 1,
        legendPos: "b",
        valAxisMinVal: axis.min,
        valAxisMaxVal: axis.max,
        // Gridlines on the axis's own round step. Left to the engine they fell
        // on half-units of the data, so points read half a unit high.
        valAxisMajorUnit: axis.step,
        valAxisLabelFormatCode: numberFormatCode(axisDecimals(axis)),
        showValAxisTitle: !!unit,
        valAxisTitle: unit,
        valAxisTitleColor: MUTED,
        valAxisTitleFontFace: BODY_FONT,
        valAxisTitleFontSize: compact ? 8 : 9,
        // A short trend carries its figures on the points, as the house
        // decks label theirs; a long one would crowd them.
        showValue: !compact && b.categories.length <= 6 && b.series.length <= 2,
        dataLabelPosition: "t",
        dataLabelFormatCode: valueFormat,
        dataLabelFontSize: 9,
        dataLabelColor: SLATE,
      }
    );
    return;
  }

  // Bar. Long or numerous category labels crowd and rotate on a vertical
  // axis, so the same data is laid out horizontally instead — a rendering
  // decision, kept in code rather than asked of the model.
  const longest = Math.max(...b.categories.map((c) => c.length));
  const horizontal = b.series.length === 1 && (longest > 14 || b.categories.length > 8);

  // A horizontal bar chart draws its first category at the BOTTOM, which
  // reverses the order the source gave. Reversing the data puts it back.
  const cats = horizontal ? [...b.categories].reverse() : b.categories;
  const series = b.series.map((s) => ({
    name: s.name,
    labels: cats,
    values: horizontal ? [...s.values].reverse() : s.values,
  }));

  // Includes zero, and reaches below it when any value is negative - a
  // negative bar used to fall below an axis pinned at 0 and simply vanish.
  const axis = niceAxis(all, { includeZero: true, pad: 0.1 });

  slide.addChart("bar", series, {
    ...common,
    barDir: horizontal ? "bar" : "col",
    // A single series must get a single colour. Handed more than one,
    // pptxgenjs switches to per-bar colouring and paints one metric across
    // five months alternating blue and orange, as if the months were
    // different categories. Confirmed by rendering.
    chartColors: series.length === 1 ? [FM_BLUE] : chartPalette(series.length),
    barGapWidthPct: 55,
    // No data labels beside a table of the same figures: three series in half
    // a slide put each label behind its neighbour's bar, so they read as
    // ",265,000". The table carries the exact values anyway.
    showValue: !compact,
    dataLabelColor: SLATE,
    dataLabelFontSize: 9,
    // A bar encodes its value as a length, so it has to start at zero.
    // Left to itself the axis baselines just under the smallest value and
    // "14 of 16 completed" renders as a near-empty bar beside a full one.
    valAxisMinVal: axis.min,
    valAxisMaxVal: axis.max,
    valAxisMajorUnit: axis.step,
    valAxisLabelFormatCode: numberFormatCode(axisDecimals(axis)),
    showValAxisTitle: !!unit,
    valAxisTitle: unit,
    valAxisTitleColor: MUTED,
    valAxisTitleFontFace: BODY_FONT,
    valAxisTitleFontSize: compact ? 8 : 9,
    showLegend: series.length > 1,
    legendPos: "b",
  });
}

// Status words a source writes into a cell, coloured as the reference's
// tracker tables colour them: the word stays, the fill only marks it. Only a
// whole-cell match counts, and no status is ever inferred - a date is never
// turned into "expired" here.
const STATUS_FILLS: { test: RegExp; fill: string; color: string }[] = [
  { test: /^(completed?|done|closed|ok|okay|working|valid|available|resolved|approved|active|renewed)$/i, fill: "E4F2E7", color: "1E6B34" },
  { test: /^(pending|open|in progress|wip|ongoing|scheduled|under process|due)$/i, fill: "FDF0D9", color: "8A5A00" },
  { test: /^(overdue|expired|not done|not working|failed|rejected|breakdown|out of order)$/i, fill: "F9E1E1", color: "9B2C2C" },
];

/** Category band rows: a pale navy tint, so they no longer read as a second header. */
const BAND_FILL = "E8EDF4";

function renderTable(
  slide: pptxgen.Slide,
  b: Extract<Block, { type: "table" }>,
  y: number,
  x: number = CONTENT_X,
  w: number = CONTENT_W
) {
  const geo = tableGeometry(b, w);
  const type = tableType(b);
  const numeric = numericColumns(b);
  const ids = identifierColumns(b);
  const align = (c: number) => (numeric[c] ? ("right" as const) : ("left" as const));

  const headerRow = b.headers.map((h, c) => ({
    text: h,
    options: {
      fill: { color: FM_BLUE }, color: WHITE, bold: true, fontSize: type.header, fontFace: BODY_FONT,
      align: align(c), valign: "middle" as const,
    },
  }));

  const body: pptxgen.TableRow[] = b.rows.map((row) =>
    row.map((cell, c) => {
      const raw = String(cell ?? "");
      const status = STATUS_FILLS.find((s) => s.test.test(raw.trim()));
      return {
        // Figures regrouped for display only (10,03,500); identifiers as typed.
        text: numeric[c] && !ids[c] ? indianFigures(raw) : raw,
        options: {
          fontSize: type.body, fontFace: BODY_FONT, color: status?.color ?? SLATE, align: align(c), valign: "middle" as const,
          ...(status ? { fill: { color: status.fill }, bold: true } : {}),
        },
      };
    })
  );

  // Insert group band rows at their stated positions, walking from the
  // bottom up so earlier insertions don't shift later indexes.
  const groups = [...(b.groups ?? [])].sort((a, b2) => b2.afterRow - a.afterRow);
  for (const g of groups) {
    const band: pptxgen.TableRow = [
      {
        text: g.label,
        options: {
          fill: { color: BAND_FILL }, color: FM_BLUE, bold: true, fontSize: type.body,
          fontFace: BODY_FONT, colspan: b.headers.length,
        },
      },
    ];
    const at = Math.max(0, Math.min(g.afterRow, body.length));
    body.splice(at, 0, band);
  }

  // A roomy table is drawn at the row heights the layout measured, so its
  // extra padding is real; otherwise rows size to their text.
  let rowH: number[] | undefined;
  if (b.roomy) {
    const bodyHs = [...geo.rowHs];
    const bandsAt = [...(b.groups ?? [])]
      .map((g, k) => ({ at: Math.max(0, Math.min(g.afterRow, b.rows.length)), h: geo.bandHs[k] }))
      .sort((p, q) => q.at - p.at);
    for (const band of bandsAt) bodyHs.splice(band.at, 0, band.h);
    rowH = [geo.headerH, ...bodyHs];
  }

  slide.addTable([headerRow, ...body], {
    x, y, w, colW: geo.widths, autoPage: false,
    ...(rowH ? { rowH } : {}),
    border: { type: "solid", color: HAIRLINE, pt: 0.5 },
  });

  if (b.footnote) {
    const fh = footnoteHeight(b.footnote, w);
    slide.addText(b.footnote, {
      x, y: y + geo.total + 0.04, w, h: fh, margin: [0, 2, 0, 2],
      fontFace: BODY_FONT, fontSize: FOOTNOTE_SIZE, italic: true, color: MUTED, valign: "top",
    });
  }
}


/**
 * Where derived copies of an image go. Session photos are already in this
 * invocation's temp folder; a bundled asset is not, and the deployment's own
 * folder is read-only on Vercel, so its copies go to the temp folder too.
 */
function workDirFor(srcPath: string): string {
  const tmp = path.resolve(os.tmpdir());
  return path.resolve(srcPath).startsWith(tmp) ? path.dirname(srcPath) : tmp;
}

/**
 * Centre-crops an image to an exact aspect ratio and returns the new path.
 *
 * This exists because pptxgenjs's `sizing` option cannot be relied on: when
 * `w` and `h` are both supplied it is ignored outright and the image is
 * simply stretched to fill the box. Verified by rendering a known circle
 * through every combination - only a box matching the image's own ratio
 * came out undistorted. So the image is made to match the box instead.
 */
async function cropToRatio(srcPath: string, ratio: number, tag: string, attention = false): Promise<string> {
  const out = path.join(workDirFor(srcPath), `crop-${tag}-${path.basename(srcPath)}`);
  if (fs.existsSync(out)) return out;
  try {
    // Never enlarged: a small photo blown up to 1400px only looked sharper
    // in the file size. It is cropped at its own resolution, capped at 1600px.
    const meta = await sharp(srcPath).metadata();
    const croppable = Math.min(meta.width ?? 1600, (meta.height ?? 1600) * ratio);
    const width = Math.max(1, Math.round(Math.min(1600, croppable)));
    await sharp(srcPath)
      .resize({
        width,
        height: Math.max(1, Math.round(width / ratio)),
        fit: "cover",
        position: attention ? sharp.strategy.attention : "centre",
      })
      .jpeg({ quality: 82 })
      .toFile(out);
    return out;
  } catch {
    return srcPath; // if cropping fails, better a slightly-off image than none
  }
}

/** Black-and-white copy of an image, for the closing slide. Falls back to the
 *  original rather than losing the photograph if conversion fails. */
async function toGrayscale(srcPath: string): Promise<string> {
  const out = path.join(workDirFor(srcPath), `bw-${path.basename(srcPath)}`);
  if (fs.existsSync(out)) return out;
  try {
    await sharp(srcPath).grayscale().jpeg({ quality: 82 }).toFile(out);
    return out;
  } catch {
    return srcPath;
  }
}

/**
 * Lays out a photo grid inside `available` vertical inches, returning the
 * height used. Photos keep their true proportions: the tile ratio is fixed,
 * each photo is centre-cropped to it, and if the grid won't fit the whole
 * thing scales down rather than squashing.
 */
async function renderPhotos(
  slide: pptxgen.Slide,
  imageIds: string[],
  x: number,
  w: number,
  y: number,
  geo: PhotoGeometry,
  sessionDir: string,
  images: ImageRecord[]
): Promise<void> {
  const found = imageIds
    .map((id) => images.find((im) => im.id === id))
    .filter((im): im is ImageRecord => !!im)
    .filter((im) => fs.existsSync(path.join(sessionDir, im.file)))
    .slice(0, 6);

  const gap = 0.18;
  const gridW = geo.cols * geo.w + (geo.cols - 1) * gap;
  const startX = x + (w - gridW) / 2; // centre the grid in its band

  for (let i = 0; i < found.length; i++) {
    const im = found[i];
    const col = i % geo.cols;
    const row = Math.floor(i / geo.cols);
    // Cropped around what draws the eye - the person, the equipment - not
    // the geometric centre, which on a portrait shot cut into a landscape
    // tile was often a wall.
    const cropped = await cropToRatio(
      path.join(sessionDir, im.file),
      geo.w / geo.h,
      `tile${(geo.w / geo.h).toFixed(3)}`,
      true
    );
    slide.addImage({
      path: cropped,
      x: startX + col * (geo.w + gap),
      y: y + row * (geo.h + gap),
      w: geo.w,
      h: geo.h,
    });
  }
}

/** Draws one block at a given position and size. */
async function renderBlock(
  slide: pptxgen.Slide,
  block: Block,
  x: number,
  y: number,
  w: number,
  h: number,
  geo: PhotoGeometry | undefined,
  sessionDir: string,
  images: ImageRecord[],
  compact = false
) {
  switch (block.type) {
    case "kpis":
      renderKpis(slide, block.items, x, y, w, w < CONTENT_W / 2);
      break;
    case "chart":
      renderChart(slide, block, y, h, x, w, compact);
      break;
    case "table":
      renderTable(slide, block, y, x, w);
      break;
    case "narrative":
      slide.addText(block.text, {
        x, y, w, h,
        fontFace: BODY_FONT, fontSize: 12.5, color: SLATE, valign: "top", lineSpacingMultiple: 1.1,
      });
      break;
    case "bullets":
      slide.addText(
        block.items.map((t) => ({
          text: t,
          options: { bullet: { characterCode: "2022" }, breakLine: true },
        })),
        { x: x + 0.1, y, w: w - 0.1, h, fontFace: BODY_FONT, fontSize: 12, color: SLATE, valign: "top", lineSpacingMultiple: 1.1 }
      );
      break;
    case "photos":
      if (geo) await renderPhotos(slide, block.imageIds, x, w, y, geo, sessionDir, images);
      break;
    case "note":
      // Dashed orange outline - the recurring container motif in the house
      // decks, and it stops a caveat reading as a stray line of text.
      slide.addShape("rect", {
        x, y, w, h,
        fill: { color: WHITE },
        line: { color: SUNSHINE, width: 1, dashType: "dash" },
      });
      slide.addText(block.text, {
        x: x + 0.16, y: y + 0.04, w: w - 0.32, h: h - 0.08,
        fontFace: BODY_FONT, fontSize: 10, color: SLATE, valign: "middle",
      });
      break;
  }
}

/** Draws a slide's rows where lib/deck-layout's placeRows puts them. */
async function renderRows(
  slide: pptxgen.Slide,
  rows: Row[],
  sessionDir: string,
  images: ImageRecord[]
) {
  for (const { row, y, h, geo } of placeRows(rows)) {
    if (row.kind === "full") {
      if (row.block.type === "photos" && !geo) {
        // The planner reserved this grid's minimum, so this should not
        // happen; if it does, say so rather than lose the photographs quietly.
        console.warn(`[deck] photo grid of ${row.block.imageIds.length} not drawn: ${h.toFixed(2)}in available`);
      }
      await renderBlock(slide, row.block, CONTENT_X, y, CONTENT_W, h, geo, sessionDir, images);
      continue;
    }
    const rightW = splitRightW(row);
    await renderBlock(slide, row.left, CONTENT_X, y, row.leftW, h, undefined, sessionDir, images, row.right.type === "table");
    await renderBlock(
      slide, row.right, CONTENT_X + row.leftW + SPLIT_GUTTER, y, rightW,
      // A table beside KPIs keeps its own height; a chart takes the band.
      row.right.type === "chart" ? h : Math.max(blockHeight(row.right, rightW), h),
      undefined, sessionDir, images, row.left.type === "chart"
    );
  }
}


export async function generateDeck(
  deck: CuratedDeck,
  sessionDir: string,
  images: ImageRecord[],
  siteName: string,
  reportMonth: string
): Promise<Buffer> {
  await prepareLogo();

  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: SLIDE_W, height: 7.5 });
  pres.layout = "WIDE";

  // ---- Cover ----
  // The reference cover: a black-and-white photograph inset under the band,
  // with the orange block and the navy name strip over its top-left corner
  // and the colour logo on the photograph's pale upper edge. The site's own
  // photo is used when it is sharp enough for the panel; otherwise SILA's
  // building from the official template, rather than a blurred enlargement
  // or a white panel.
  const coverPhoto = chooseCoverPhoto(deck, images, sessionDir);
  const panelPhoto = await toGrayscale(
    await cropToRatio(
      coverPhoto ? path.join(sessionDir, coverPhoto.file) : fallbackBuilding(),
      PHOTO_PANEL.w / PHOTO_PANEL.h,
      "panel"
    )
  );

  const cover = pres.addSlide();
  cover.background = { color: WHITE };
  band(cover);
  cover.addImage({ path: panelPhoto, ...PHOTO_PANEL });
  // From just inside the band, so no hairline of photo shows between them.
  cover.addShape("rect", { x: 0, y: 0.74, w: 5.63, h: 3.01, fill: { color: SUNSHINE }, line: { type: "none" } });
  cover.addShape("rect", { x: 0.76, y: 3.01, w: 4.87, h: 0.74, fill: { color: FM_BLUE }, line: { type: "none" } });
  {
    // The colour logo on the photograph's upper edge. SILA's building has a
    // pale sky there; a site photo can be dark anywhere, so it gets a white
    // plate to keep the logo legible.
    const w = 1.9;
    const h = w / COLOR_LOGO_RATIO;
    if (coverPhoto) {
      cover.addShape("rect", { x: 5.63, y: 0.75, w: w + 0.3, h: h + 0.3, fill: { color: WHITE }, line: { type: "none" } });
    }
    cover.addImage({ path: brandLogo(false), x: 5.78, y: 0.9, w, h });
  }
  cover.addText("MONTHLY MANAGEMENT REPORT", {
    x: 0.76, y: 2.35, w: 4.8, h: 0.4, margin: 0,
    fontFace: BODY_FONT, fontSize: 12, color: CHARCOAL, charSpacing: 4,
  });
  cover.addText(siteName, {
    x: 0.96, y: 3.01, w: 2.9, h: 0.74, margin: 0, valign: "middle", fit: "shrink",
    fontFace: HEAD_FONT, fontSize: 18, bold: true, color: WHITE, charSpacing: 2,
  });
  cover.addText("www.silagroup.co.in", {
    x: 3.85, y: 3.01, w: 1.65, h: 0.74, margin: 0, valign: "middle", align: "right",
    fontFace: BODY_FONT, fontSize: 9.5, color: WHITE, underline: { style: "sng" },
  });
  cover.addText(prettyMonth(reportMonth), {
    x: 0.76, y: 1.2, w: 4.8, h: 0.5, margin: 0,
    fontFace: HEAD_FONT, fontSize: 20, color: CHARCOAL,
  });
  pageSquare(cover, 1);

  // ---- At a glance ----
  let pageNum = 2;
  const glance = pres.addSlide();
  chrome(glance, "This Month at a Glance", pageNum++);
  glance.addText(deck.summary.headline, {
    x: CONTENT_X, y: CONTENT_TOP, w: CONTENT_W, h: 0.85,
    fontFace: HEAD_FONT, fontSize: 17, color: FM_BLUE, valign: "top",
  });
  glance.addShape("rect", { x: CONTENT_X, y: CONTENT_TOP + 0.95, w: CONTENT_W, h: 0.02, fill: { color: HAIRLINE }, line: { type: "none" } });
  deck.summary.points.slice(0, 5).forEach((pt, i) => {
    const y = CONTENT_TOP + 1.2 + i * 0.62;
    glance.addShape("rect", { x: CONTENT_X, y: y + 0.08, w: 0.09, h: 0.32, fill: { color: SUNSHINE }, line: { type: "none" } });
    glance.addText(pt, {
      x: CONTENT_X + 0.28, y, w: CONTENT_W - 0.28, h: 0.5,
      fontFace: BODY_FONT, fontSize: 12.5, color: SLATE, valign: "middle",
    });
  });

  // ---- Contents ----
  // The house decks open with one, and it is the slide that tells a client
  // what the month actually covered before they start paging through it.
  const contents = pres.addSlide();
  chrome(contents, `Contents — ${prettyMonth(reportMonth)}`, pageNum++);
  {
    const labels = deck.sections.map((s) => s.label);
    const cols = labels.length > 8 ? 2 : 1;
    const perCol = Math.ceil(labels.length / cols);
    const colW = (CONTENT_W - 0.6) / cols;
    labels.forEach((label, i) => {
      const col = Math.floor(i / perCol);
      const row = i % perCol;
      const x = CONTENT_X + col * (colW + 0.6);
      const y = CONTENT_TOP + 0.1 + row * 0.46;
      slide_contentsRow(contents, label, x, y, colW);
    });
  }

  // ---- One slide per parameter, continued onto another where it overflows ----
  for (const section of deck.sections) {
    // No placeholder branch: Claude omits parameters with no real content,
    // so anything reaching here has genuine blocks to render. A shorter
    // deck of real substance beats a fixed structure full of holes.
    for (const page of sectionSlides(section)) {
      const slide = pres.addSlide();
      chrome(slide, page.label, pageNum++);
      await renderRows(slide, page.rows, sessionDir, images);
    }
  }

  // ---- Thank you ----
  // The reference closing slide: the same black-and-white photograph as the
  // cover, the white logo on it top-right, an orange block bottom-left with
  // the sign-off, and the navy strip carrying the web address. Bookending
  // with the cover photo is deliberate: it is the shot chosen to show the
  // property, not "some other photo".
  const closing = pres.addSlide();
  closing.background = { color: WHITE };
  band(closing);
  closing.addImage({ path: panelPhoto, ...PHOTO_PANEL });
  {
    // The reference puts the white logo on a dark part of its photo. Where the
    // photo is not known to be dark there, the colour logo is the one that
    // reads: on SILA's building the white one vanished into the sky.
    const w = 2.0;
    if (coverPhoto) {
      closing.addShape("rect", { x: 12.55 - w - 0.15, y: 0.75, w: w + 0.3, h: w / COLOR_LOGO_RATIO + 0.3, fill: { color: WHITE }, line: { type: "none" } });
    }
    closing.addImage({ path: brandLogo(false), x: 12.55 - w, y: 0.9, w, h: w / COLOR_LOGO_RATIO });
  }
  closing.addShape("rect", { x: 0, y: 4.5, w: 5.93, h: 2.27, fill: { color: SUNSHINE }, line: { type: "none" } });
  closing.addText("THANK YOU", {
    x: 0.76, y: 4.5, w: 5.0, h: 2.27, margin: 0, valign: "middle",
    fontFace: HEAD_FONT, fontSize: 36, color: CHARCOAL, charSpacing: 6,
  });
  closing.addShape("rect", { x: 0.76, y: 6.77, w: 5.17, h: 0.73, fill: { color: FM_BLUE }, line: { type: "none" } });
  closing.addText("www.silagroup.co.in", {
    x: 0.76, y: 6.77, w: 5.17, h: 0.73, margin: 0, align: "center", valign: "middle",
    fontFace: BODY_FONT, fontSize: 10.5, color: WHITE,
  });
  pageSquare(closing, pageNum);

  const buf = await pres.write({ outputType: "nodebuffer" });
  return buf as Buffer;
}