import * as XLSX from "xlsx";

function isRealRow(row: unknown[]): boolean {
  const meaningful = row.filter((c) => c !== null && c !== undefined && c !== "" && c !== 0);
  if (meaningful.length === 0) return false;
  if (meaningful.every((c) => typeof c === "string" && c.startsWith("#"))) return false;
  return true;
}

export function extractXlsx(buffer: Buffer): { sheets: Record<string, unknown[][]>; summary: string } {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheets: Record<string, unknown[][]> = {};

  for (const sheetName of wb.SheetNames) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 }) as unknown[][];
    sheets[sheetName] = raw.filter(isRealRow);
  }

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
