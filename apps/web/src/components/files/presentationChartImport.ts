import type { PresentationChart } from "./presentationDeck";

function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "result" in value) return cellText(value.result);
  if (typeof value === "object" && "text" in value) return String(value.text);
  return String(value);
}

function chartFromRows(rows: unknown[][], source: string): PresentationChart {
  if (rows.length < 2 || rows[0]!.length < 2)
    throw new Error("Use a header row, a label column, and at least one numeric column.");
  const headers = rows[0]!
    .slice(1, 5)
    .map((value, index) => cellText(value) || `Series ${index + 1}`);
  const data = rows.slice(1, 102).filter((row) => cellText(row[0]).trim());
  if (!data.length) throw new Error("No labeled data rows were found.");
  const categories = data.map((row) => cellText(row[0]));
  const series = headers.map((name, index) => ({
    name,
    values: data.map((row) => {
      const value = row[index + 1];
      const number =
        typeof value === "number" ? value : Number(cellText(value).replaceAll(",", ""));
      return Number.isFinite(number) ? number : 0;
    }),
  }));
  if (!series.some((item) => item.values.some((value) => value !== 0)))
    throw new Error("No numeric chart values were found.");
  return { type: "bar", categories, series, source };
}

function parseDelimited(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < text.length; index++) {
    const char = text[index]!;
    if (char === '"') {
      if (quoted && text[index + 1] === '"') {
        value += '"';
        index++;
      } else quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && text[index + 1] === "\n") index++;
      row.push(value);
      if (row.some((cell) => cell.trim())) rows.push(row);
      row = [];
      value = "";
    } else value += char;
  }
  row.push(value);
  if (row.some((cell) => cell.trim())) rows.push(row);
  return rows;
}

export async function importChart(file: File): Promise<PresentationChart> {
  if (file.size > 8 * 1024 * 1024) throw new Error("Choose a spreadsheet under 8 MB.");
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv"))
    return chartFromRows(
      parseDelimited(await file.text(), name.endsWith(".tsv") ? "\t" : ","),
      file.name,
    );
  if (!name.endsWith(".xlsx")) throw new Error("Choose an .xlsx, .csv, or .tsv file.");
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("The workbook has no worksheets.");
  const rows: unknown[][] = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    if (rows.length >= 102) return;
    rows.push(
      Array.from(
        { length: Math.min(row.cellCount, 5) },
        (_, index) => row.getCell(index + 1).value,
      ),
    );
  });
  return chartFromRows(rows, `${file.name} · ${sheet.name}`);
}
