import pptxgen from "pptxgenjs";
import sharp from "sharp";
import fs from "fs";
import os from "os";
import path from "path";
import type { CuratedDeck, Block } from "./curate";
import { axisDecimals, chartPalette, decimalsNeeded, niceAxis, numberFormatCode } from "./chart-style";
import {
  BLOCK_SPACING,
  CHART_MAX_H,
  CHART_MIN_H,
  CONTENT_BOTTOM,
  CONTENT_TOP,
  CONTENT_W,
  SPLIT_GUTTER,
  SPLIT_W,
  type PhotoGeometry,
  type Row,
  photoGeometry,
  photosMinHeight,
  rowIsFlexible,
  rowMinHeight,
  sectionSlides,
} from "./deck-layout";

// FM service-line palette. The parent SILA guide assigns Blue #264170 to
// Facility Management and pairs service colours with Sunshine rather than
// Charcoal, so this is the correct palette for an FM deliverable.
const FM_BLUE = "264170";
const SUNSHINE = "F7A328";
const SLATE = "2D2D2D";
const MUTED = "697784";
const HAIRLINE = "E5E5E4";
const WHITE = "FFFFFF";

// Brand fonts are Merriweather (headings) / Avenir Next (body). Neither is
// reliably installed on a Windows machine, and a missing font renders as an
// ugly substitute, so these are deliberate near-equivalents: Georgia is a
// serif in the same register as Merriweather, Calibri a clean sans for body.
const HEAD_FONT = "Georgia";
const BODY_FONT = "Calibri";

const SLIDE_W = 13.333;
// Card height for a KPI row; deck-layout costs the block at the same figure.
const KPI_H = 1.15;
const CONTENT_X = 0.6;

type ImageRecord = { id: string; file: string; slideNumber: number; slideTitle: string };


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
 * Downscaled copy of the white logo, prepared once per build.
 *
 * pptxgenjs embeds an image per addImage call rather than once per file, so the
 * full-size 180KB logo was stored 18 times - 3.3MB of one picture, more than
 * half the deck. It is only ever drawn 1.42in wide.
 */
let preparedLogo: string | null = null;
async function prepareLogo(): Promise<void> {
  const src = logo(true);
  const out = path.join(os.tmpdir(), "sila-fm-logo-white-320.png");
  try {
    if (!fs.existsSync(out)) await sharp(src).resize({ width: 320 }).png().toFile(out);
    preparedLogo = out;
  } catch {
    preparedLogo = null; // fall back to the full-size original
  }
}
function brandLogo(): string {
  return preparedLogo ?? logo(true);
}

function chrome(slide: pptxgen.Slide, title: string, pageNum: number) {
  slide.background = { color: WHITE };
  slide.addShape("rect", { x: 0, y: 0, w: SLIDE_W, h: 0.65, fill: { color: FM_BLUE }, line: { type: "none" } });

  slide.addImage({ path: brandLogo(), x: 11.62, y: 0.13, w: 1.42, h: 0.4 });

  // Set at the scale the house decks use. At 26pt the heading read as a
  // caption rather than a title and left the slide looking unanchored; the
  // reference decks run their titles around twice the body size again.
  slide.addText(title, {
    x: CONTENT_X, y: 0.74, w: 10.4, h: 0.72,
    fontFace: HEAD_FONT, fontSize: 34, color: SLATE, valign: "middle",
  });
  slide.addShape("rect", { x: CONTENT_X, y: 1.52, w: 1.5, h: 0.05, fill: { color: SUNSHINE }, line: { type: "none" } });

  slide.addShape("rect", { x: 12.52, y: 6.92, w: 0.58, h: 0.42, fill: { color: SUNSHINE }, line: { type: "none" } });
  slide.addText(String(pageNum), {
    x: 12.52, y: 6.92, w: 0.58, h: 0.42, align: "center", valign: "middle",
    fontFace: BODY_FONT, fontSize: 12, bold: true, color: WHITE,
  });
}

/** One line of the contents list: orange tick bar, then the section name. */
function slide_contentsRow(slide: pptxgen.Slide, label: string, x: number, y: number, w: number) {
  slide.addShape("rect", { x, y: y + 0.08, w: 0.1, h: 0.22, fill: { color: SUNSHINE }, line: { type: "none" } });
  slide.addText(label, {
    x: x + 0.3, y, w: w - 0.3, h: 0.38,
    fontFace: BODY_FONT, fontSize: 13, color: SLATE, valign: "middle",
  });
}

function renderKpis(slide: pptxgen.Slide, items: { label: string; value: string }[], y: number) {
  // Carded like the house decks: a bordered panel with an orange edge and
  // the figure set large. The brand guide calls for stat callouts far bigger
  // than body copy, and at 22pt in a flat grey box these read as footnotes.
  const shown = items.slice(0, 4);
  const gap = 0.22;
  const w = (CONTENT_W - gap * (shown.length - 1)) / shown.length;
  shown.forEach((kpi, i) => {
    const x = CONTENT_X + i * (w + gap);
    slide.addShape("rect", {
      x, y, w, h: KPI_H,
      fill: { color: WHITE },
      line: { color: HAIRLINE, width: 1 },
    });
    slide.addShape("rect", { x, y, w: 0.07, h: KPI_H, fill: { color: SUNSHINE }, line: { type: "none" } });
    slide.addText(kpi.value, {
      x: x + 0.24, y: y + 0.12, w: w - 0.36, h: 0.6,
      fontFace: HEAD_FONT, fontSize: 30, color: FM_BLUE, valign: "middle", shrinkText: true,
    });
    slide.addText(kpi.label.toUpperCase(), {
      x: x + 0.26, y: y + 0.74, w: w - 0.38, h: 0.3,
      fontFace: BODY_FONT, fontSize: 9, color: MUTED, charSpacing: 1.2, valign: "middle",
    });
  });
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

  const common = {
    x,
    y,
    w,
    h,
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
        showValue: false,
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

function renderTable(
  slide: pptxgen.Slide,
  b: Extract<Block, { type: "table" }>,
  y: number,
  x: number = CONTENT_X,
  w: number = CONTENT_W
) {
  const headerRow = b.headers.map((h) => ({
    text: h,
    options: { fill: { color: FM_BLUE }, color: WHITE, bold: true, fontSize: 10.5, fontFace: BODY_FONT },
  }));

  // Insert group band rows at their stated positions, walking from the
  // bottom up so earlier insertions don't shift later indexes.
  const body: pptxgen.TableRow[] = b.rows.map((row) =>
    row.map((cell) => ({
      text: String(cell ?? ""),
      options: { fontSize: 10, fontFace: BODY_FONT, color: SLATE },
    }))
  );

  const groups = [...(b.groups ?? [])].sort((a, b2) => b2.afterRow - a.afterRow);
  for (const g of groups) {
    const band: pptxgen.TableRow = [
      {
        text: g.label,
        options: {
          fill: { color: FM_BLUE }, color: WHITE, bold: true, fontSize: 10,
          fontFace: BODY_FONT, colspan: b.headers.length,
        },
      },
    ];
    const at = Math.max(0, Math.min(g.afterRow, body.length));
    body.splice(at, 0, band);
  }

  slide.addTable([headerRow, ...body], {
    x, y, w, autoPage: false,
    border: { type: "solid", color: HAIRLINE, pt: 0.5 },
  });
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
async function cropToRatio(srcPath: string, ratio: number, tag: string): Promise<string> {
  const dir = path.dirname(srcPath);
  const out = path.join(dir, `crop-${tag}-${path.basename(srcPath)}`);
  if (fs.existsSync(out)) return out;
  try {
    const width = 1400;
    await sharp(srcPath)
      .resize({ width, height: Math.round(width / ratio), fit: "cover", position: "centre" })
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
  const out = path.join(path.dirname(srcPath), `bw-${path.basename(srcPath)}`);
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
  const startX = CONTENT_X + (CONTENT_W - gridW) / 2; // centre the grid

  for (let i = 0; i < found.length; i++) {
    const im = found[i];
    const col = i % geo.cols;
    const row = Math.floor(i / geo.cols);
    const cropped = await cropToRatio(
      path.join(sessionDir, im.file),
      geo.w / geo.h,
      `${geo.cols}x${geo.rows}`
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

/** Draws one block at a given position and width. */
async function renderBlock(
  slide: pptxgen.Slide,
  block: Block,
  y: number,
  h: number,
  x: number,
  w: number,
  geo: PhotoGeometry | undefined,
  sessionDir: string,
  images: ImageRecord[],
  compact = false
) {
  switch (block.type) {
    case "kpis":
      renderKpis(slide, block.items, y);
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
        fontFace: BODY_FONT, fontSize: 12.5, color: SLATE, valign: "top",
      });
      break;
    case "bullets":
      slide.addText(
        block.items.map((t) => ({
          text: t,
          options: { bullet: { characterCode: "2022" }, breakLine: true },
        })),
        { x: x + 0.1, y, w: w - 0.1, h, fontFace: BODY_FONT, fontSize: 12, color: SLATE, valign: "top" }
      );
      break;
    case "photos":
      if (geo) await renderPhotos(slide, block.imageIds, y, geo, sessionDir, images);
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

/**
 * Lays rows out in two passes: measure everything first, then start low enough
 * that the whole group sits centred in the content area. Stacking straight from
 * the top left slides looking bottom-heavy with a dead band underneath -
 * measuring first is what makes them read as composed.
 */
async function renderRows(
  slide: pptxgen.Slide,
  rows: Row[],
  sessionDir: string,
  images: ImageRecord[]
) {
  // Same constant the planner measured with, so the two cannot disagree about
  // what fits.
  const spacing = BLOCK_SPACING;
  const areaH = CONTENT_BOTTOM - CONTENT_TOP;

  type Measured = { row: Row; h: number; leftGeo?: PhotoGeometry; geo?: PhotoGeometry };
  const measured: Measured[] = [];
  let fixedTotal = 0;

  for (const row of rows) {
    if (rowIsFlexible(row)) {
      measured.push({ row, h: 0 });
    } else {
      const h = rowMinHeight(row);
      fixedTotal += h;
      measured.push({ row, h });
    }
  }
  const gaps = spacing * Math.max(0, rows.length - 1);
  let flexBudget = areaH - fixedTotal - gaps;

  // Flexible rows take their minimum first, then share out whatever is left.
  // Sizing a split row to its minimum was the bug preflight found next: a
  // three-row table set the band at an inch and a half, so the chart beside it
  // stayed a strip and its data labels wrapped mid-number. A split row should
  // fill the space it has.
  const isChartRow = (m: Measured) => m.row.kind === "full" && m.row.block.type === "chart";
  const isPhotoRow = (m: Measured) => m.row.kind === "full" && m.row.block.type === "photos";
  const isSplitRow = (m: Measured) => m.row.kind === "split";

  const growable = measured.filter((m) => isSplitRow(m) || isChartRow(m));
  const photoBlock = (m: Measured) =>
    (m.row as { kind: "full"; block: Block }).block as Extract<Block, { type: "photos" }>;
  // Each grid's real minimum, so a chart growing into the spare space cannot
  // leave a grid less room than it needs to be drawn at all.
  const photoReserve = measured
    .filter(isPhotoRow)
    .reduce((sum, m) => sum + photosMinHeight(photoBlock(m).imageIds.length), 0);

  for (const m of growable) {
    m.h = Math.max(CHART_MIN_H, rowMinHeight(m.row));
    flexBudget -= m.h;
  }

  // Share the remainder, capped per kind: a chart on its own should not
  // balloon across a whole slide, but a chart paired with its table should
  // take the full band.
  let spare = Math.max(0, flexBudget - photoReserve);
  for (let i = 0; i < growable.length && spare > 0.01; i++) {
    const m = growable[i];
    const ceiling = isSplitRow(m) ? areaH : CHART_MAX_H;
    const room = Math.max(0, ceiling - m.h);
    const give = Math.min(room, spare / (growable.length - i));
    m.h += give;
    spare -= give;
    flexBudget -= give;
  }

  // Photo grids share what is left: each first gets its own minimum, then the
  // excess is split between them. The first grid used to be handed the whole
  // remainder, so a second grid on the same slide got nothing and was not
  // drawn; an equal split is no better when the grids differ, since a
  // six-photo grid needs more room than a three-photo one.
  const photoRows = measured.filter(isPhotoRow);
  let excess = flexBudget - photoReserve;
  let photoRowsLeft = photoRows.length;
  for (const m of photoRows) {
    const block = photoBlock(m);
    const min = photosMinHeight(block.imageIds.length);
    const share = min + Math.max(0, excess) / Math.max(1, photoRowsLeft);
    photoRowsLeft--;
    const geo = photoGeometry(Math.min(block.imageIds.length, 6), share);
    if (geo) {
      m.geo = geo;
      m.h = geo.used;
      flexBudget -= geo.used;
      excess -= Math.max(0, geo.used - min);
    } else {
      // The planner reserved this grid's minimum, so this should not happen;
      // if it does, say so rather than lose the photographs quietly.
      console.warn(`[deck] photo grid of ${block.imageIds.length} not drawn: ${share.toFixed(2)}in available`);
    }
  }

  const total = measured.reduce((sum, m) => sum + m.h, 0) + gaps;
  let y = CONTENT_TOP + Math.max(0, (areaH - total) / 2);

  for (const m of measured) {
    if (m.h === 0 && rowIsFlexible(m.row)) continue;

    if (m.row.kind === "full") {
      await renderBlock(slide, m.row.block, y, m.h, CONTENT_X, CONTENT_W, m.geo, sessionDir, images);
    } else {
      await renderBlock(slide, m.row.left, y, m.h, CONTENT_X, SPLIT_W, undefined, sessionDir, images, true);
      await renderBlock(
        slide, m.row.right, y, m.h,
        CONTENT_X + SPLIT_W + SPLIT_GUTTER, SPLIT_W,
        undefined, sessionDir, images, true
      );
    }
    y += m.h + spacing;
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
  const cover = pres.addSlide();
  cover.background = { color: WHITE };
  cover.addShape("rect", { x: 0, y: 0, w: SLIDE_W, h: 0.65, fill: { color: FM_BLUE }, line: { type: "none" } });

  // Right ~60%: full-bleed photo. Claude nominates one (it knows which
  // source slide each came from, so it can tell a building exterior from
  // a photo of a battery rack); largest file is only the fallback, and it
  // picked badly often enough to be worth replacing.
  const nominated = deck.coverImageId
    ? images.find((im) => im.id === deck.coverImageId)
    : undefined;
  const coverPhoto =
    (nominated && fs.existsSync(path.join(sessionDir, nominated.file)) ? nominated : undefined) ??
    [...images]
      .filter((im) => fs.existsSync(path.join(sessionDir, im.file)))
      .sort(
        (a, b) =>
          fs.statSync(path.join(sessionDir, b.file)).size -
          fs.statSync(path.join(sessionDir, a.file)).size
      )[0];

  if (coverPhoto) {
    const w = SLIDE_W - 5.3;
    const h = 6.85;
    // Cropped to the panel's ratio for the same reason as the grid tiles:
    // a mismatched box stretches the image rather than cropping it.
    const cropped = await cropToRatio(path.join(sessionDir, coverPhoto.file), w / h, "cover");
    cover.addImage({ path: cropped, x: 5.3, y: 0.65, w, h });
  }

  cover.addShape("rect", { x: 0, y: 0.65, w: 5.3, h: 6.85, fill: { color: SUNSHINE }, line: { type: "none" } });
  cover.addShape("rect", { x: 0, y: 3.15, w: 5.3, h: 1.5, fill: { color: FM_BLUE }, line: { type: "none" } });

  cover.addText("MONTHLY MANAGEMENT REPORT", {
    x: 0.45, y: 1.5, w: 4.5, h: 0.5,
    fontFace: BODY_FONT, fontSize: 13, color: SLATE, charSpacing: 3, bold: true,
  });
  cover.addText(siteName, {
    x: 0.45, y: 3.3, w: 4.5, h: 1.2,
    fontFace: HEAD_FONT, fontSize: 30, color: WHITE, valign: "middle",
  });
  cover.addText(prettyMonth(reportMonth), {
    x: 0.45, y: 4.85, w: 4.5, h: 0.4,
    fontFace: BODY_FONT, fontSize: 14, color: SLATE,
  });
  cover.addText("www.silagroup.co.in", {
    x: 0.45, y: 6.65, w: 4.5, h: 0.35,
    fontFace: BODY_FONT, fontSize: 11, color: SLATE,
  });

  // In the header bar rather than over the photo: a white logo on a light
  // photo was unreadable, and the photo's brightness can't be predicted.
  cover.addImage({ path: brandLogo(), x: 11.62, y: 0.13, w: 1.42, h: 0.4 });

  // ---- At a glance ----
  let pageNum = 2;
  const glance = pres.addSlide();
  chrome(glance, "This Month at a Glance", pageNum++);
  glance.addText(deck.summary.headline, {
    x: CONTENT_X, y: CONTENT_TOP, w: CONTENT_W, h: 0.85,
    fontFace: HEAD_FONT, fontSize: 17, color: FM_BLUE, valign: "top",
  });
  glance.addShape("rect", { x: CONTENT_X, y: 2.85, w: CONTENT_W, h: 0.02, fill: { color: HAIRLINE }, line: { type: "none" } });
  deck.summary.points.slice(0, 5).forEach((pt, i) => {
    const y = 3.1 + i * 0.62;
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
  // Required by the brand guide and present in every house deck: a full-bleed
  // photograph with an orange band across it carrying the sign-off. The photo
  // is a second site image where one exists, so the deck closes on the
  // property rather than on a blank panel.
  const closing = pres.addSlide();
  closing.background = { color: WHITE };
  closing.addShape("rect", { x: 0, y: 0, w: SLIDE_W, h: 0.65, fill: { color: FM_BLUE }, line: { type: "none" } });

  // The cover photograph again, in black and white, bookending the deck the
  // way the house decks do. Deliberately not "some other photo": Claude picks
  // the cover knowing which slide each image came from, so it is the one shot
  // established to show the property. Choosing by file size instead landed on
  // a before/after composite with "After" captioned across it.
  const closingPhoto = coverPhoto;
  if (closingPhoto) {
    const w = SLIDE_W - 1.6;
    const h = 5.6;
    const cropped = await cropToRatio(path.join(sessionDir, closingPhoto.file), w / h, "closing");
    closing.addImage({ path: await toGrayscale(cropped), x: 0.8, y: 0.65, w, h });
  }

  closing.addShape("rect", { x: 0, y: 3.55, w: 7.4, h: 1.75, fill: { color: SUNSHINE }, line: { type: "none" } });
  closing.addText("THANK YOU", {
    x: 0.8, y: 3.55, w: 6.2, h: 1.75,
    fontFace: HEAD_FONT, fontSize: 40, color: SLATE, charSpacing: 4, valign: "middle",
  });
  closing.addShape("rect", { x: 0, y: 6.6, w: 6.2, h: 0.9, fill: { color: FM_BLUE }, line: { type: "none" } });
  closing.addText("www.silagroup.co.in", {
    x: 0.8, y: 6.6, w: 5, h: 0.9,
    fontFace: BODY_FONT, fontSize: 13, color: WHITE, valign: "middle",
  });
  closing.addImage({ path: brandLogo(), x: 11.62, y: 0.13, w: 1.42, h: 0.4 });

  const buf = await pres.write({ outputType: "nodebuffer" });
  return buf as Buffer;
}