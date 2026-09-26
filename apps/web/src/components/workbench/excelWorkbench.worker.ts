import ExcelJS from "exceljs";
import JSZip from "jszip";

import {
  WORKBOOK_MAX_COLUMNS,
  WORKBOOK_MAX_CELLS,
  WORKBOOK_MAX_FILE_BYTES,
  WORKBOOK_MAX_ROWS,
  WORKBOOK_PAGE_SIZE,
  type WorkbookCell,
  type WorkbookChartPoint,
  type WorkbookColumnProfile,
  type WorkbookCommand,
  type WorkbookRequest,
  type WorkbookResponse,
  type WorkbookSheetOverview,
} from "./excelWorkbenchModel";

type WorkerScope = {
  onmessage: ((event: MessageEvent<WorkbookRequest>) => void) | null;
  postMessage: (message: WorkbookResponse, transfer?: Transferable[]) => void;
};

interface ActiveSheet {
  readonly name: string;
  readonly columns: ReadonlyArray<string>;
  readonly rows: ReadonlyArray<ReadonlyArray<WorkbookCell>>;
  readonly overview: WorkbookSheetOverview;
}

interface XlsxTableRange {
  readonly firstColumn: number;
  readonly firstRow: number;
  readonly lastColumn: number;
  readonly lastRow: number;
}

const workerScope = self as unknown as WorkerScope;
let workbook: ExcelJS.Workbook | null = null;
let csvSheets: ReadonlyArray<ActiveSheet> = [];
let activeSheet: ActiveSheet | null = null;
let xlsxTableRanges: ReadonlyMap<string, XlsxTableRange> = new Map();
let workbookFileName = "customer-data.xlsx";
let workbookFileSize = 0;
let sheetNames: ReadonlyArray<string> = [];

workerScope.onmessage = (event) => {
  void handleRequest(event.data);
};

async function handleRequest(request: WorkbookRequest): Promise<void> {
  try {
    switch (request.type) {
      case "load": {
        const loaded = await loadWorkbook(request.file);
        activeSheet = loaded.firstSheet;
        sheetNames = loaded.sheetNames;
        workbookFileName = request.file.name;
        workbookFileSize = request.file.size;
        sendLoaded(
          request.requestId,
          "loaded",
          loaded.firstSheet,
          0,
          loaded.firstSheet.rows.length,
        );
        break;
      }
      case "sheet": {
        const nextSheet = workbook
          ? makeXlsxSheet(workbook.getWorksheet(request.name), xlsxTableRanges.get(request.name))
          : csvSheets.find((sheet) => sheet.name === request.name);
        if (!nextSheet) throw new Error("That worksheet could not be opened.");
        activeSheet = nextSheet;
        sendLoaded(request.requestId, "sheet", nextSheet, 0, nextSheet.rows.length);
        break;
      }
      case "query": {
        const sheet = requireActiveSheet();
        const matchedRows = findRows(sheet, request.term);
        const start = request.page * WORKBOOK_PAGE_SIZE;
        workerScope.postMessage({
          type: "rows",
          requestId: request.requestId,
          rows: matchedRows.slice(start, start + WORKBOOK_PAGE_SIZE),
          page: request.page,
          filteredRowCount: matchedRows.length,
        });
        break;
      }
      case "ask": {
        workerScope.postMessage({
          type: "answer",
          requestId: request.requestId,
          text: answerQuestion(requireActiveSheet(), request.question),
        });
        break;
      }
      case "chart": {
        workerScope.postMessage({
          type: "chart",
          requestId: request.requestId,
          points: buildCategoryTotals(
            requireActiveSheet(),
            request.groupColumn,
            request.valueColumn,
          ),
        });
        break;
      }
      case "context": {
        const sheets = workbook
          ? workbook.worksheets.map((sheet) =>
              makeXlsxSheet(sheet, xlsxTableRanges.get(sheet.name)),
            )
          : csvSheets;
        const sections = sheets.map((sheet) => {
          const columns = sheet.columns.join("\t");
          const rows = sheet.rows
            .slice(0, 1000)
            .map((row) =>
              row.map((cell) => String(cell ?? "").replaceAll(/[\t\r\n]+/g, " ")).join("\t"),
            );
          return `## ${sheet.name}\n${sheet.overview.rowCount} rows; ${sheet.overview.duplicateRows} duplicate rows; ${sheet.overview.missingCells} missing cells.\n${columns}\n${rows.join("\n")}${sheet.rows.length > 1000 ? "\n[Only the first 1000 rows are included in AI context.]" : ""}`;
        });
        workerScope.postMessage({
          type: "context",
          requestId: request.requestId,
          text: sections.join("\n\n").slice(0, 800_000),
        });
        break;
      }
      case "export": {
        const sheet = requireActiveSheet();
        const data = await createAnalysisWorkbook(sheet);
        workerScope.postMessage(
          {
            type: "exported",
            requestId: request.requestId,
            fileName: buildExportName(workbookFileName),
            data,
          },
          [data],
        );
        break;
      }
    }
  } catch (error) {
    workerScope.postMessage({
      type: "error",
      requestId: request.requestId,
      message: error instanceof Error ? error.message : "The workbook could not be processed.",
    });
  }
}

async function loadWorkbook(
  file: File,
): Promise<{ firstSheet: ActiveSheet; sheetNames: ReadonlyArray<string> }> {
  if (file.size > WORKBOOK_MAX_FILE_BYTES) {
    throw new Error("Choose a workbook smaller than 20 MB.");
  }

  const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  if (extension === ".csv" || extension === ".tsv") {
    const text = await file.text();
    const delimiter = extension === ".tsv" ? "\t" : detectCsvDelimiter(text);
    const nextCsvSheets = parseDelimitedFile(text, delimiter);
    const firstSheet = nextCsvSheets[0];
    if (!firstSheet) throw new Error("This file does not contain data.");
    workbook = null;
    xlsxTableRanges = new Map();
    csvSheets = nextCsvSheets;
    return { firstSheet, sheetNames: nextCsvSheets.map((sheet) => sheet.name) };
  }
  if (extension !== ".xlsx") {
    throw new Error("Use an .xlsx, .csv, or .tsv file. Legacy .xls files are not supported yet.");
  }

  const normalizedXlsx = await normalizeXlsx(file);
  const nextWorkbook = new ExcelJS.Workbook();
  await nextWorkbook.xlsx.load(normalizedXlsx.data);
  const firstWorksheet = nextWorkbook.worksheets[0];
  if (!firstWorksheet) throw new Error("This workbook does not contain any worksheets.");
  const firstSheet = makeXlsxSheet(
    firstWorksheet,
    normalizedXlsx.tableRanges.get(firstWorksheet.name),
  );
  workbook = nextWorkbook;
  xlsxTableRanges = normalizedXlsx.tableRanges;
  csvSheets = [];
  return { firstSheet, sheetNames: nextWorkbook.worksheets.map((worksheet) => worksheet.name) };
}

async function normalizeXlsx(
  file: File,
): Promise<{ data: ArrayBuffer; tableRanges: ReadonlyMap<string, XlsxTableRange> }> {
  const archive = await JSZip.loadAsync(await file.arrayBuffer());
  const tableRanges = await findWorksheetTableRanges(archive);

  for (const [name, entry] of Object.entries(archive.files)) {
    if (entry.dir) continue;
    if (name.startsWith("xl/tables/") || name.startsWith("xl/drawings/")) {
      archive.remove(name);
      continue;
    }
    if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(name)) {
      const relationships = await entry.async("string");
      archive.file(
        name,
        relationships.replace(/<(?:[\w.-]+:)?Relationship\b[^>]*\/>/g, (relationship) => {
          const type = readXmlAttribute(relationship, "Type") ?? "";
          return type.endsWith("/table") || type.endsWith("/drawing") ? "" : relationship;
        }),
      );
      continue;
    }
    if (!name.endsWith(".xml")) continue;

    let xml = await entry.async("string");
    xml = xml.replace(/([<\/])x:/g, "$1").replace(/\sxmlns:x=/g, " xmlns=");
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      xml = xml
        .replace(/<tableParts\b[^>]*>[\s\S]*?<\/tableParts>/g, "")
        .replace(/<tableParts\b[^>]*\/>/g, "")
        .replace(/<drawing\b[^>]*\/>/g, "")
        .replace(/<drawing\b[^>]*>[\s\S]*?<\/drawing>/g, "");
    }
    if (name === "[Content_Types].xml") {
      xml = xml.replace(/<Default\b[^>]*\/>/g, (defaultType) => {
        if (readXmlAttribute(defaultType, "Extension") !== "xml") return defaultType;
        return defaultType.replace(/\bContentType="[^"]*"/, 'ContentType="application/xml"');
      });
      xml = xml.replace(/<Override\b[^>]*\/>/g, (override) => {
        const partName = readXmlAttribute(override, "PartName") ?? "";
        return /^\/xl\/(?:tables|drawings)\//.test(partName) ? "" : override;
      });
      if (!/<Override\b[^>]*PartName="\/xl\/workbook\.xml"/.test(xml)) {
        xml = xml.replace(
          "</Types>",
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /></Types>',
        );
      }
    }
    archive.file(name, xml);
  }

  return {
    data: await archive.generateAsync({ type: "arraybuffer" }),
    tableRanges,
  };
}

async function findWorksheetTableRanges(
  archive: JSZip,
): Promise<ReadonlyMap<string, XlsxTableRange>> {
  const workbookXml = await readZipText(archive, "xl/workbook.xml");
  const workbookRelationships = await readZipText(archive, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !workbookRelationships) return new Map();

  const workbookRelationshipById = new Map(
    readXmlStartTags(workbookRelationships, "Relationship").flatMap((attributes) => {
      const id = readXmlAttribute(attributes, "Id");
      const target = readXmlAttribute(attributes, "Target");
      return id && target ? [[id, target] as const] : [];
    }),
  );
  const ranges = new Map<string, XlsxTableRange>();

  for (const attributes of readXmlStartTags(workbookXml, "sheet")) {
    const name = readXmlAttribute(attributes, "name");
    const relationshipId = readXmlAttribute(attributes, "r:id");
    const sheetTarget = relationshipId ? workbookRelationshipById.get(relationshipId) : undefined;
    if (!name || !sheetTarget) continue;

    const sheetPath = resolvePackagePath("xl/workbook.xml", sheetTarget);
    const sheetRelationships = await readZipText(archive, relationshipPath(sheetPath));
    if (!sheetRelationships) continue;

    const sheetRanges: Array<XlsxTableRange> = [];
    for (const attributes of readXmlStartTags(sheetRelationships, "Relationship")) {
      const type = readXmlAttribute(attributes, "Type") ?? "";
      const target = readXmlAttribute(attributes, "Target");
      if (!type.endsWith("/table") || !target) continue;
      const tablePath = resolvePackagePath(sheetPath, target);
      const tableXml = await readZipText(archive, tablePath);
      const tableAttributes = tableXml ? readXmlStartTags(tableXml, "table")[0] : undefined;
      const reference = tableAttributes ? readXmlAttribute(tableAttributes, "ref") : undefined;
      const range = reference ? parseTableRange(reference) : null;
      if (range) sheetRanges.push(range);
    }

    sheetRanges.sort(
      (left, right) => left.firstRow - right.firstRow || left.firstColumn - right.firstColumn,
    );
    const firstRange = sheetRanges[0];
    if (firstRange) ranges.set(decodeXmlText(name), firstRange);
  }
  return ranges;
}

async function readZipText(archive: JSZip, path: string): Promise<string | null> {
  const entry = archive.file(path);
  return entry ? entry.async("string") : null;
}

function readXmlStartTags(xml: string, tagName: string): Array<string> {
  const expression = new RegExp(`<(?:(?:[\\w.-]+):)?${tagName}\\b([^>]*)>`, "g");
  return [...xml.matchAll(expression)].map((match) => match[1] ?? "");
}

function readXmlAttribute(attributes: string, name: string): string | undefined {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return attributes.match(new RegExp(`(?:^|\\s)${escapedName}="([^"]*)"`))?.[1];
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function resolvePackagePath(basePath: string, target: string): string {
  const segments = target.startsWith("/") ? [] : basePath.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
}

function relationshipPath(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  const directory = separator < 0 ? "" : partPath.slice(0, separator);
  const fileName = partPath.slice(separator + 1);
  return `${directory}/_rels/${fileName}.rels`;
}

function parseTableRange(reference: string): XlsxTableRange | null {
  const match = reference.replace(/\$/g, "").match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return null;
  const firstColumn = columnNumber(match[1] ?? "");
  const firstRow = Number(match[2]);
  const lastColumn = columnNumber(match[3] ?? "");
  const lastRow = Number(match[4]);
  if (firstColumn < 1 || firstRow < 1 || lastColumn < firstColumn || lastRow < firstRow) {
    return null;
  }
  return { firstColumn, firstRow, lastColumn, lastRow };
}

function columnNumber(letters: string): number {
  return [...letters.toUpperCase()].reduce(
    (value, letter) => value * 26 + letter.charCodeAt(0) - 64,
    0,
  );
}

function makeXlsxSheet(
  worksheet: ExcelJS.Worksheet | undefined,
  tableRange?: XlsxTableRange,
): ActiveSheet {
  if (!worksheet) throw new Error("That worksheet could not be opened.");
  const rawRows: Array<ReadonlyArray<WorkbookCell>> = [];
  let maxColumnCount = 0;

  if (tableRange) {
    const rowCount = tableRange.lastRow - tableRange.firstRow + 1;
    const columnCount = tableRange.lastColumn - tableRange.firstColumn + 1;
    if (columnCount > WORKBOOK_MAX_COLUMNS) {
      throw new Error(`Worksheets can have up to ${String(WORKBOOK_MAX_COLUMNS)} columns.`);
    }
    if (rowCount > WORKBOOK_MAX_ROWS + 1) {
      throw new Error(
        `Worksheets can have up to ${String(WORKBOOK_MAX_ROWS.toLocaleString())} data rows.`,
      );
    }
    if (rowCount * columnCount > WORKBOOK_MAX_CELLS) {
      throw new Error(
        "This worksheet is too large to analyze in the browser. Filter it to fewer than 1.5 million cells and try again.",
      );
    }
    for (let rowNumber = tableRange.firstRow; rowNumber <= tableRange.lastRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      rawRows.push(
        Array.from({ length: columnCount }, (_, index) =>
          normalizeExcelValue(row.getCell(tableRange.firstColumn + index).value),
        ),
      );
    }
    maxColumnCount = columnCount;
  } else {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      let width = 0;
      row.eachCell({ includeEmpty: false }, (_cell, columnNumber) => {
        width = Math.max(width, columnNumber);
      });
      if (width > WORKBOOK_MAX_COLUMNS) {
        throw new Error(`Worksheets can have up to ${String(WORKBOOK_MAX_COLUMNS)} columns.`);
      }
      const values = Array.from({ length: width }, (_, index) =>
        normalizeExcelValue(row.getCell(index + 1).value),
      );
      while (values.length > 0 && isBlank(values[values.length - 1])) values.pop();
      if (values.length > 0 && !values.every(isBlank)) {
        rawRows.push(values);
        maxColumnCount = Math.max(maxColumnCount, values.length);
        if (rawRows.length * maxColumnCount > WORKBOOK_MAX_CELLS) {
          throw new Error(
            "This worksheet is too large to analyze in the browser. Filter it to fewer than 1.5 million cells and try again.",
          );
        }
        if (rawRows.length > WORKBOOK_MAX_ROWS + 1) {
          throw new Error(
            `Worksheets can have up to ${String(WORKBOOK_MAX_ROWS.toLocaleString())} data rows.`,
          );
        }
      }
    });
    const headerCandidate = rawRows.findIndex(
      (row) => row.filter((value) => !isBlank(value)).length > 1,
    );
    if (
      headerCandidate > 0 &&
      rawRows
        .slice(0, headerCandidate)
        .every((row) => row.filter((value) => !isBlank(value)).length === 1)
    ) {
      rawRows.splice(0, headerCandidate);
      maxColumnCount = rawRows.reduce((largest, row) => Math.max(largest, row.length), 0);
    }
  }
  return makeSheet(worksheet.name, rawRows, maxColumnCount);
}

function parseDelimitedFile(text: string, delimiter: "," | "\t" | ";"): ReadonlyArray<ActiveSheet> {
  const matrix = parseDelimitedText(text, delimiter);
  const rows = matrix.filter((row) => !row.every(isBlank));
  const width = rows.reduce((largest, row) => Math.max(largest, row.length), 0);
  if (rows.length * width > WORKBOOK_MAX_CELLS) {
    throw new Error(
      "This file is too large to analyze in the browser. Filter it to fewer than 1.5 million cells and try again.",
    );
  }
  return [makeSheet("Data", rows, width)];
}

function parseDelimitedText(text: string, delimiter: "," | "\t" | ";"): Array<Array<WorkbookCell>> {
  const rows: Array<Array<WorkbookCell>> = [];
  let row: Array<WorkbookCell> = [];
  let cell = "";
  let quoted = false;
  let cellCount = 0;
  const source = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;

  const pushCell = () => {
    row.push(parseDelimitedValue(cell, delimiter));
    cell = "";
    cellCount++;
    if (cellCount > WORKBOOK_MAX_CELLS) {
      throw new Error(
        "This file is too large to analyze in the browser. Filter it to fewer than 1.5 million cells and try again.",
      );
    }
    if (row.length > WORKBOOK_MAX_COLUMNS) {
      throw new Error(`Worksheets can have up to ${String(WORKBOOK_MAX_COLUMNS)} columns.`);
    }
  };

  for (let index = 0; index < source.length; index++) {
    const character = source[index];
    if (character === '"') {
      if (quoted && source[index + 1] === '"') {
        cell += '"';
        index++;
      } else if (quoted || cell.length === 0) {
        quoted = !quoted;
      } else {
        cell += character;
      }
      continue;
    }
    if (!quoted && (character === delimiter || character === "\r" || character === "\n")) {
      pushCell();
      if (character !== delimiter) {
        rows.push(row);
        row = [];
        if (character === "\r" && source[index + 1] === "\n") index++;
        if (rows.length > WORKBOOK_MAX_ROWS + 1) {
          throw new Error(
            `Files can have up to ${String(WORKBOOK_MAX_ROWS.toLocaleString())} data rows.`,
          );
        }
      }
      continue;
    }
    cell += character;
  }
  if (quoted) throw new Error("This delimited file has an unclosed quoted cell.");
  if (cell.length > 0 || row.length > 0) {
    pushCell();
    rows.push(row);
  }
  return rows;
}

function parseDelimitedValue(value: string, delimiter: "," | "\t" | ";"): WorkbookCell {
  const trimmed = value.trim();
  if (trimmed === "") return null;
  if (delimiter === ";" && /^-?(?:[1-9]\d{0,2}(?:\.\d{3})+|0|[1-9]\d*)(?:,\d+)?$/.test(trimmed)) {
    const number = Number(trimmed.replace(/\./g, "").replace(",", "."));
    if (Number.isFinite(number)) return number;
  }
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(trimmed)) {
    const number = Number(trimmed);
    if (Number.isFinite(number)) return number;
  }
  return value;
}

function detectCsvDelimiter(text: string): "," | ";" {
  let commas = 0;
  let semicolons = 0;
  let quoted = false;
  for (const character of text.slice(0, 16_384)) {
    if (character === '"') quoted = !quoted;
    if (quoted) continue;
    if (character === "\n" || character === "\r") break;
    if (character === ",") commas++;
    if (character === ";") semicolons++;
  }
  return semicolons > commas ? ";" : ",";
}

function makeSheet(
  name: string,
  rows: ReadonlyArray<ReadonlyArray<WorkbookCell>>,
  columnCount: number,
): ActiveSheet {
  const headerRow = rows[0];
  const columns = makeHeaders(headerRow ?? [], columnCount);
  const tableRows = rows
    .slice(1)
    .map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? null));
  const overview = analyzeSheet(name, columns, tableRows);
  return { name, columns, rows: tableRows, overview };
}

function makeHeaders(
  values: ReadonlyArray<WorkbookCell>,
  columnCount: number,
): ReadonlyArray<string> {
  const seen = new Map<string, number>();
  return Array.from({ length: columnCount }, (_, index) => {
    const base = formatCell(values[index]).trim() || `Column ${String(index + 1)}`;
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base} (${String(count)})`;
  });
}

function analyzeSheet(
  name: string,
  columns: ReadonlyArray<string>,
  rows: ReadonlyArray<ReadonlyArray<WorkbookCell>>,
): WorkbookSheetOverview {
  const accumulators = columns.map(() => ({
    values: new Set<string>(),
    types: new Set<WorkbookColumnProfile["type"]>(),
    examples: [] as Array<string>,
    numericCount: 0,
    numericTotal: 0,
    numericMinimum: Number.POSITIVE_INFINITY,
    numericMaximum: Number.NEGATIVE_INFINITY,
    missing: 0,
  }));
  let missingCells = 0;
  let duplicateRows = 0;
  const seenRows = new Set<string>();

  for (const row of rows) {
    const normalized = Array.from({ length: columns.length }, (_, index) => row[index] ?? null);
    const rowSignature = JSON.stringify(normalized);
    if (seenRows.has(rowSignature)) duplicateRows++;
    else seenRows.add(rowSignature);

    for (let index = 0; index < columns.length; index++) {
      const value = normalized[index] ?? null;
      const accumulator = accumulators[index];
      if (!accumulator) continue;
      if (isBlank(value)) {
        accumulator.missing++;
        missingCells++;
        continue;
      }
      const display = formatCell(value);
      accumulator.values.add(display);
      if (accumulator.examples.length < 2 && !accumulator.examples.includes(display)) {
        accumulator.examples.push(display);
      }
      if (typeof value === "number") {
        accumulator.types.add("Number");
        accumulator.numericCount++;
        accumulator.numericTotal += value;
        accumulator.numericMinimum = Math.min(accumulator.numericMinimum, value);
        accumulator.numericMaximum = Math.max(accumulator.numericMaximum, value);
      } else if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(value)) {
        accumulator.types.add("Date");
      } else {
        accumulator.types.add("Text");
      }
    }
  }

  const profiles = columns.map((column, index): WorkbookColumnProfile => {
    const accumulator = accumulators[index];
    const numericCount = accumulator?.numericCount ?? 0;
    const total = accumulator?.numericTotal ?? 0;
    const kinds = accumulator?.types ?? new Set<WorkbookColumnProfile["type"]>();
    const type = kinds.size === 0 ? "Text" : kinds.size === 1 ? [...kinds][0]! : "Mixed";
    const hasNumericValues = type === "Number" && numericCount > 0;
    return {
      name: column,
      type,
      nonEmpty: rows.length - (accumulator?.missing ?? 0),
      missing: accumulator?.missing ?? 0,
      unique: accumulator?.values.size ?? 0,
      examples: accumulator?.examples ?? [],
      total: hasNumericValues ? total : null,
      average: hasNumericValues ? total / numericCount : null,
      minimum: hasNumericValues ? (accumulator?.numericMinimum ?? null) : null,
      maximum: hasNumericValues ? (accumulator?.numericMaximum ?? null) : null,
    };
  });
  const cellCount = rows.length * columns.length;
  return {
    name,
    rowCount: rows.length,
    columns: profiles,
    duplicateRows,
    missingCells,
    completeness:
      cellCount === 0 ? 100 : Math.round(((cellCount - missingCells) / cellCount) * 100),
  };
}

function findRows(sheet: ActiveSheet, term: string): ReadonlyArray<ReadonlyArray<WorkbookCell>> {
  const normalizedTerm = term.trim().toLocaleLowerCase();
  if (!normalizedTerm) return sheet.rows;
  return sheet.rows.filter((row) =>
    row.some((value) => formatCell(value).toLocaleLowerCase().includes(normalizedTerm)),
  );
}

function answerQuestion(sheet: ActiveSheet, question: string): string {
  const normalizedQuestion = question.trim().toLocaleLowerCase();
  if (!normalizedQuestion) return "Enter a question about this worksheet.";

  if (/\b(duplicate|duplicates|repeated|dubbel|dubbele)\b/.test(normalizedQuestion)) {
    return `${formatCell(sheet.overview.duplicateRows)} duplicate record${sheet.overview.duplicateRows === 1 ? "" : "s"} found.`;
  }
  if (/\b(missing|empty|blank|null|ontbrekend|ontbrekende|leeg|lege)\b/.test(normalizedQuestion)) {
    const fields = sheet.overview.columns
      .filter((column) => column.missing > 0)
      .sort((left, right) => right.missing - left.missing);
    if (fields.length === 0) return "No empty cells found in this worksheet.";
    const fieldSummary = fields
      .slice(0, 5)
      .map((column) => `${column.name} (${formatCell(column.missing)})`)
      .join(", ");
    return `${formatCell(sheet.overview.missingCells)} empty cells across ${formatCell(sheet.overview.rowCount)} records. Fields with the most gaps: ${fieldSummary}${fields.length > 5 ? ", and more" : ""}.`;
  }
  if (
    /\b(how many|count|number of|records|customers|rows|hoeveel|aantal|rijen)\b/.test(
      normalizedQuestion,
    )
  ) {
    return `This worksheet has ${formatCell(sheet.overview.rowCount)} rows across ${formatCell(sheet.overview.columns.length)} columns.`;
  }

  const column = [...sheet.overview.columns]
    .sort((left, right) => right.name.length - left.name.length)
    .find((candidate) => normalizedQuestion.includes(candidate.name.toLocaleLowerCase()));
  if (column?.type === "Number" && column.average !== null) {
    if (/\b(average|avg|mean|gemiddelde)\b/.test(normalizedQuestion)) {
      return `The average ${column.name} is ${formatNumber(column.average)}.`;
    }
    if (/\b(total|sum|totaal|som)\b/.test(normalizedQuestion)) {
      return `The total ${column.name} is ${formatNumber(column.total ?? 0)}.`;
    }
    if (/\b(minimum|lowest|smallest|min|laagste|kleinste)\b/.test(normalizedQuestion)) {
      return `The lowest ${column.name} is ${formatNumber(column.minimum ?? 0)}.`;
    }
    if (/\b(maximum|highest|largest|max|hoogste|grootste)\b/.test(normalizedQuestion)) {
      return `The highest ${column.name} is ${formatNumber(column.maximum ?? 0)}.`;
    }
  }

  return "Try asking for the record count, missing values, duplicates, or the total, average, minimum, or maximum of a numeric field.";
}

function buildCategoryTotals(
  sheet: ActiveSheet,
  groupColumn: string,
  valueColumn: string,
): ReadonlyArray<WorkbookChartPoint> {
  const groupIndex = sheet.columns.indexOf(groupColumn);
  const valueIndex = sheet.columns.indexOf(valueColumn);
  if (groupIndex < 0 || valueIndex < 0) return [];

  const totals = new Map<string, { total: number; records: number }>();
  for (const row of sheet.rows) {
    const groupValue = row[groupIndex];
    const numericValue = row[valueIndex];
    if (
      (typeof groupValue !== "string" &&
        typeof groupValue !== "number" &&
        typeof groupValue !== "boolean") ||
      typeof numericValue !== "number"
    ) {
      continue;
    }
    const label = formatCell(groupValue).trim();
    if (!label) continue;
    const entry = totals.get(label) ?? { total: 0, records: 0 };
    entry.total += numericValue;
    entry.records++;
    totals.set(label, entry);
  }

  return [...totals]
    .map(([label, value]) => ({ label, ...value }))
    .sort((left, right) => Math.abs(right.total) - Math.abs(left.total))
    .slice(0, 8);
}

function formatNumber(value: number): string {
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(value);
}

function sendLoaded(
  requestId: number,
  type: "loaded" | "sheet",
  sheet: ActiveSheet,
  page: number,
  filteredRowCount: number,
): void {
  workerScope.postMessage({
    type,
    requestId,
    fileName: workbookFileName,
    fileSize: workbookFileSize,
    sheetNames,
    overview: sheet.overview,
    rows: sheet.rows.slice(0, WORKBOOK_PAGE_SIZE),
    page,
    filteredRowCount,
  });
}

function requireActiveSheet(): ActiveSheet {
  if (!activeSheet) throw new Error("Upload a workbook to get started.");
  return activeSheet;
}

function normalizeExcelValue(value: unknown): WorkbookCell {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if ("result" in value) {
      if (value.result !== undefined) return normalizeExcelValue(value.result);
      if ("formula" in value && typeof value.formula === "string") return `=${value.formula}`;
      return null;
    }
    if ("text" in value && typeof value.text === "string") return value.text;
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText
        .map((part) =>
          typeof part === "object" &&
          part !== null &&
          "text" in part &&
          typeof part.text === "string"
            ? part.text
            : "",
        )
        .join("");
    }
  }
  return String(value);
}

function formatCell(value: WorkbookCell | undefined): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  return String(value);
}

function isBlank(value: WorkbookCell | undefined): boolean {
  return (
    value === null || value === undefined || (typeof value === "string" && value.trim() === "")
  );
}

async function createAnalysisWorkbook(sheet: ActiveSheet): Promise<ArrayBuffer> {
  const output = new ExcelJS.Workbook();
  output.creator = "T3 Code";
  output.subject = "Customer data analysis";
  output.title = `${sheet.name} analysis`;
  const dataName = cleanSheetName(sheet.name);
  const data = output.addWorksheet(
    dataName.toLowerCase() === "analysis" ? "Source data" : dataName,
  );
  data.addRow([...sheet.columns]);
  for (const row of sheet.rows) data.addRow([...row]);
  data.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  data.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253041" } };
  data.views = [{ state: "frozen", ySplit: 1 }];
  data.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(1, sheet.rows.length + 1), column: Math.max(1, sheet.columns.length) },
  };
  for (let index = 1; index <= sheet.columns.length; index++) {
    const header = sheet.columns[index - 1] ?? "";
    data.getColumn(index).width = Math.min(36, Math.max(12, header.length + 3));
  }

  const analysis = output.addWorksheet("Analysis");
  analysis.addRows([
    ["Workbook analysis", ""],
    ["Worksheet", sheet.name],
    ["Rows", sheet.overview.rowCount],
    ["Columns", sheet.columns.length],
    ["Completeness", `${String(sheet.overview.completeness)}%`],
    ["Duplicate rows", sheet.overview.duplicateRows],
    [],
    ["Column", "Type", "Non-empty", "Missing", "Unique", "Total", "Average", "Minimum", "Maximum"],
    ...sheet.overview.columns.map((column) => [
      column.name,
      column.type,
      column.nonEmpty,
      column.missing,
      column.unique,
      column.total,
      column.average,
      column.minimum,
      column.maximum,
    ]),
  ]);
  analysis.getRow(1).font = { bold: true, size: 14 };
  analysis.getRow(8).font = { bold: true, color: { argb: "FFFFFFFF" } };
  analysis.getRow(8).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF253041" } };
  analysis.views = [{ state: "frozen", ySplit: 8 }];
  analysis.columns.forEach((column, index) => {
    column.width = index === 0 ? 28 : 16;
  });

  const bytes = await output.xlsx.writeBuffer();
  const result = new Uint8Array(bytes);
  return result.buffer.slice(
    result.byteOffset,
    result.byteOffset + result.byteLength,
  ) as ArrayBuffer;
}

function cleanSheetName(name: string): string {
  const normalized = name.replace(/[\\/?*:[\]]/g, " ").trim();
  return (normalized || "Data").slice(0, 31);
}

function buildExportName(name: string): string {
  const base = name.replace(/\.[^.]+$/, "").trim() || "customer-data";
  return `${base}-analysis.xlsx`;
}
