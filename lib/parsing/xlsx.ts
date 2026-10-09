import * as XLSX from "xlsx";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * Renders a date-formatted numeric cell as text the model cannot misread,
 * or returns null to leave the cell's number untouched.
 *
 * Excel stores a date as a serial day count, so without this a Month column
 * reaches the model as 45748 rather than April 2025. A real MMR had its
 * electricity months stored as April-June 2025 under slides labelled
 * Jun-26..Aug-26, and the model paired the 2025 figures with the 2026 labels
 * because it could not read the serials. Excel's own display text is no
 * better: "Apr-25" reads equally well as 25 April, so the text is rebuilt
 * from the date itself, always with a four-digit year.
 */
function dateCellText(cell: XLSX.CellObject, date1904: boolean): string | null {
  if (cell.t !== "n" || typeof cell.v !== "number" || typeof cell.z !== "string") return null;
  if (!XLSX.SSF.is_date(cell.z)) return null;
  // A zero in a date column is almost always an empty formula result, and
  // isRealRow drops rows on exactly that zero, so it stays a number.
  if (cell.v <= 0) return null;

  // Elapsed-time formats ([h]:mm) are durations, not dates; Excel's own
  // rendering of those is already unambiguous.
  if (/\[(h+|m+|s+)\]/i.test(cell.z)) return cell.w ?? null;
  // Only the positive section of the format matters, and quoted literals,
  // [$-409]-style locale tags and backslash escapes are not date tokens.
  const tokens = cell.z.split(";")[0].replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");
  const hasTime = /[hs]/i.test(tokens) || /am\/pm|a\/p/i.test(tokens);
  // A lone "m" next to hours is minutes, so a format with time but no d or y
  // is a time of day, which Excel already renders unambiguously.
  const hasDate = /[dy]/i.test(tokens) || (!hasTime && /m/i.test(tokens));
  if (!hasDate) return cell.w ?? null;

  const d = XLSX.SSF.parse_date_code(cell.v, { date1904 });
  if (!d || d.m < 1 || d.m > 12) return null;
  // "Apr 2025" when the format shows no day, so a month column is not
  // mistaken for the 1st of the month; "1 Apr 2025" when it does.
  let text = /d/i.test(tokens) ? `${d.d} ${MONTHS[d.m - 1]} ${d.y}` : `${MONTHS[d.m - 1]} ${d.y}`;
  if (hasTime) text += ` ${String(d.H).padStart(2, "0")}:${String(d.M).padStart(2, "0")}`;
  return text;
}

/**
 * Reads every sheet of a workbook as rows of cell values, with date cells
 * decoded to readable text and every other value left raw. Numbers must stay
 * numbers: isRealRow's zero filter and the chart values depend on it. Shared
 * by standalone workbooks and those embedded in a PPTX so the two paths
 * cannot drift apart.
 */
export function readWorkbookRows(buffer: Buffer): Record<string, unknown[][]> {
  // cellNF keeps each cell's number format, which is the only thing that
  // marks a number as a date.
  const wb = XLSX.read(buffer, { type: "buffer", cellNF: true });
  const date1904 = !!wb.Workbook?.WBProps?.date1904;
  const sheets: Record<string, unknown[][]> = {};
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const cell = ws[addr] as XLSX.CellObject;
      const text = dateCellText(cell, date1904);
      if (text !== null) {
        cell.t = "s";
        cell.v = text;
        cell.w = text;
      }
      // Once decoded here, a cell's date format has done its job and goes.
      // SheetJS 0.20's sheet_to_json turns any number carrying a date format
      // into a Date in the server's local time - 1 Apr 2025 came out as
      // "2025-03-31T18:30:00.000Z" on an IST machine - and returned the text
      // cells written above as null. A date left numeric here (a zero from an
      // empty formula) must stay the number isRealRow filters on.
      if (typeof cell.z === "string" && XLSX.SSF.is_date(cell.z)) delete cell.z;
    }
    sheets[sheetName] = XLSX.utils.sheet_to_json(ws, { header: 1 }) as unknown[][];
  }
  return sheets;
}

function isRealRow(row: unknown[]): boolean {
  const meaningful = row.filter((c) => c !== null && c !== undefined && c !== "" && c !== 0);
  if (meaningful.length === 0) return false;
  if (meaningful.every((c) => typeof c === "string" && c.startsWith("#"))) return false;
  return true;
}

/** Rows of one sheet shown to the model before the rest are left out. A real
 *  submission's master workbooks reach 2.1M characters in full against 380k
 *  at this limit, so the limit stays and a cut sheet says so. */
export const SHEET_ROW_LIMIT = 25;

/** The line that ends a sheet the model is not shown in full. The prompt
 *  refers to it by this wording. */
export const CUT_SHORT = "... further rows not shown - this sheet is cut short here";

/**
 * A sheet as the model reads it: a heading, then its rows as JSON.
 *
 * It carries no row count. The count used to be printed as "(N real data
 * rows)", but N included the header, title and TOTAL rows, so a 40-task sheet
 * was announced as 42. Once a sheet was cut short, that wrong figure was the
 * only total the model could see, and the grounding check accepted it because
 * it was written in the source. A count that cannot be stated correctly is not
 * stated.
 */
export function sheetText(heading: string, rows: unknown[][], indent = ""): string {
  const lines = [heading];
  for (const r of rows.slice(0, SHEET_ROW_LIMIT)) lines.push(indent + JSON.stringify(r));
  if (rows.length > SHEET_ROW_LIMIT) lines.push(indent + CUT_SHORT);
  return lines.join("\n");
}

/** Every non-empty sheet of a workbook, as the model reads it. */
export function workbookText(sheets: Record<string, unknown[][]>): string {
  return Object.entries(sheets)
    .filter(([, rows]) => rows.length > 0)
    .map(([name, rows]) => sheetText(`--- Sheet: ${name} ---`, rows))
    .join("\n\n");
}

export function extractXlsx(buffer: Buffer): { sheets: Record<string, unknown[][]>; summary: string } {
  const sheets: Record<string, unknown[][]> = {};
  for (const [sheetName, raw] of Object.entries(readWorkbookRows(buffer))) {
    sheets[sheetName] = raw.filter(isRealRow);
  }

  // `summary` is what app/api/parse-submission returns to n8n, so its wording
  // is left exactly as that workflow has always received it. This app's own
  // prompts use workbookText.
  const summary = Object.entries(sheets)
    .filter(([, rows]) => rows.length > 0)
    .map(([sheetName, rows]) => {
      const parts = [`--- Sheet: ${sheetName} (${rows.length} real data rows) ---`];
      for (const row of rows.slice(0, 15)) parts.push(JSON.stringify(row));
      if (rows.length > 15) parts.push(`... ${rows.length - 15} more rows`);
      return parts.join("\n");
    })
    .join("\n\n");

  return { sheets, summary };
}
