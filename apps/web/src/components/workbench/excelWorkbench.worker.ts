import { PROVIDER_SEND_TURN_MAX_INPUT_CHARS } from "@t3tools/contracts";
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
  readonly sourceRowNumbers: ReadonlyArray<number>;
  readonly firstColumn: number;
  readonly headerRowNumber: number;
}

interface XlsxTableRange {
  readonly firstColumn: number;
  readonly firstRow: number;
  readonly lastColumn: number;
  readonly lastRow: number;
}

const workerScope = self as unknown as WorkerScope;
let workbook: ExcelJS.Workbook | null = null;
let originalXlsx: ArrayBuffer | null = null;
let xlsxEdits = new Map<string, Map<string, WorkbookCell>>();
let structuralDirty = false;
let structureWarning: string | null = null;
let sheetOverrides = new Map<string, ActiveSheet>();
let csvSheets: ReadonlyArray<ActiveSheet> = [];
let csvDelimiter: "," | "\t" | ";" = ",";
let activeSheet: ActiveSheet | null = null;
let xlsxTableRanges: ReadonlyMap<string, XlsxTableRange> = new Map();
let workbookFileName = "customer-data.xlsx";
let workbookFileSize = 0;
let sheetNames: ReadonlyArray<string> = [];
let completeAgentContextCache: string | null = null;
let completeAgentContextFailure: Error | null = null;

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
        const nextSheet =
          sheetOverrides.get(request.name) ??
          (workbook
            ? makeXlsxSheet(workbook.getWorksheet(request.name), xlsxTableRanges.get(request.name))
            : csvSheets.find((sheet) => sheet.name === request.name));
        if (!nextSheet) throw new Error("That worksheet could not be opened.");
        activeSheet = nextSheet;
        sendLoaded(request.requestId, "sheet", nextSheet, 0, nextSheet.rows.length);
        break;
      }
      case "query": {
        const sheet = requireActiveSheet();
        const matchedIndexes = findRowIndexes(sheet, request.term);
        const start = request.page * WORKBOOK_PAGE_SIZE;
        const rowIndexes = matchedIndexes.slice(start, start + WORKBOOK_PAGE_SIZE);
        workerScope.postMessage({
          type: "rows",
          requestId: request.requestId,
          rows: rowIndexes.map((index) => sheet.rows[index]!),
          rowIndexes,
          page: request.page,
          filteredRowCount: matchedIndexes.length,
        });
        break;
      }
      case "editCell": {
        const sheet = requireActiveSheet();
        if (
          !Number.isInteger(request.row) ||
          request.row < 0 ||
          request.row >= sheet.rows.length ||
          !Number.isInteger(request.column) ||
          request.column < 0 ||
          request.column >= sheet.columns.length
        ) {
          throw new Error("That cell is outside this worksheet.");
        }
        const rows = sheet.rows.map((row, index) =>
          index === request.row
            ? row.map((cell, column) => (column === request.column ? request.value : cell))
            : row,
        );
        if (workbook) {
          const worksheet = workbook.getWorksheet(sheet.name);
          if (!worksheet) throw new Error("That worksheet could not be edited.");
          const cell = worksheet
            .getRow(sheet.sourceRowNumbers[request.row]!)
            .getCell(sheet.firstColumn + request.column);
          cell.value = request.value;
          const edits = xlsxEdits.get(sheet.name) ?? new Map<string, WorkbookCell>();
          edits.set(cell.address, request.value);
          xlsxEdits.set(sheet.name, edits);
        } else {
          csvSheets = csvSheets.map((current) =>
            current.name === sheet.name
              ? { ...current, rows, overview: analyzeSheet(current.name, current.columns, rows) }
              : current,
          );
        }
        activeSheet = { ...sheet, rows, overview: analyzeSheet(sheet.name, sheet.columns, rows) };
        if (workbook && structuralDirty) sheetOverrides.set(sheet.name, activeSheet);
        completeAgentContextCache = null;
        completeAgentContextFailure = null;
        sendLoaded(request.requestId, "sheet", activeSheet, 0, rows.length);
        break;
      }
      case "insertRow":
      case "deleteRows":
      case "insertColumn":
      case "deleteColumns":
      case "addSheet": {
        const nextSheet = changeWorkbookStructure(request);
        activeSheet = nextSheet;
        completeAgentContextCache = null;
        completeAgentContextFailure = null;
        sendLoaded(request.requestId, "sheet", nextSheet, 0, nextSheet.rows.length);
        break;
      }
      case "save": {
        requireActiveSheet();
        const data = workbook
          ? structuralDirty
            ? await workbook.xlsx.writeBuffer()
            : await saveOriginalXlsx()
          : new TextEncoder().encode(
              [csvSheets[0]!.columns, ...csvSheets[0]!.rows]
                .map((row) =>
                  row.map((value) => quoteDelimitedCell(value, csvDelimiter)).join(csvDelimiter),
                )
                .join("\r\n") + "\r\n",
            ).buffer;
        const bytes = new Uint8Array(data);
        const buffer = bytes.buffer.slice(
          bytes.byteOffset,
          bytes.byteOffset + bytes.byteLength,
        ) as ArrayBuffer;
        workerScope.postMessage({ type: "saved", requestId: request.requestId, data: buffer }, [
          buffer,
        ]);
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
          ? workbook.worksheets.map(
              (sheet) =>
                sheetOverrides.get(sheet.name) ??
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
      case "agentContext": {
        requireActiveSheet();
        const text = buildCompleteAgentContext();
        workerScope.postMessage({ type: "context", requestId: request.requestId, text });
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

function buildCompleteAgentContext(): string {
  if (completeAgentContextFailure) throw new Error(completeAgentContextFailure.message);
  if (completeAgentContextCache !== null) return completeAgentContextCache;
  const maximumChars = PROVIDER_SEND_TURN_MAX_INPUT_CHARS - 2_000;
  const parts: Array<string> = [`Workbook: ${workbookFileName}\n`];
  let characterCount = parts[0]?.length ?? 0;
  const append = (text: string) => {
    characterCount += text.length;
    if (characterCount > maximumChars) {
      completeAgentContextFailure = new Error(
        "This workbook is too large to include completely in one agent message. No workbook data was sent. Split it into smaller files and try again.",
      );
      throw new Error(completeAgentContextFailure.message);
    }
    parts.push(text);
  };
  const appendSheet = (sheet: ActiveSheet) => {
    append(
      `\n## Worksheet: ${JSON.stringify(sheet.name)}\nColumns: ${JSON.stringify(sheet.columns)}\nRows are JSON arrays in their original order:\n`,
    );
    for (const row of sheet.rows) append(`${JSON.stringify(row)}\n`);
  };

  if (workbook) {
    for (const worksheet of workbook.worksheets) {
      appendSheet(
        sheetOverrides.get(worksheet.name) ??
          makeXlsxSheet(worksheet, xlsxTableRanges.get(worksheet.name)),
      );
    }
  } else {
    for (const sheet of csvSheets) appendSheet(sheet);
  }

  completeAgentContextCache = parts.join("");
  return completeAgentContextCache;
}

function changeWorkbookStructure(
  request: Extract<
    WorkbookRequest,
    { type: "insertRow" | "deleteRows" | "insertColumn" | "deleteColumns" | "addSheet" }
  >,
): ActiveSheet {
  if (structureWarning) throw new Error(structureWarning);
  if (request.type === "addSheet") {
    if (!workbook) throw new Error("CSV and TSV files have only one worksheet.");
    const name = request.name.trim();
    if (!name || name.length > 31 || /[\\/?*:[\]]/.test(name))
      throw new Error("Choose a valid worksheet name of 1 to 31 characters.");
    if (sheetNames.some((existing) => existing.toLocaleLowerCase() === name.toLocaleLowerCase()))
      throw new Error("A worksheet with that name already exists.");
    workbook.addWorksheet(name);
    sheetNames = workbook.worksheets.map((sheet) => sheet.name);
    const next = makeSheet(name, [], 0);
    sheetOverrides.set(name, next);
    structuralDirty = true;
    return next;
  }
  const sheet = requireActiveSheet();
  const worksheet = workbook?.getWorksheet(sheet.name);
  const indexes = "indexes" in request ? [...new Set(request.indexes)].sort((a, b) => b - a) : [];
  let columns = [...sheet.columns];
  let rows = sheet.rows.map((row) => [...row]);
  let sourceRowNumbers = [...sheet.sourceRowNumbers];
  if (request.type === "insertRow") {
    if (!Number.isInteger(request.index) || request.index < 0 || request.index > rows.length)
      throw new Error("Choose a valid row position.");
    if ((rows.length + 1) * columns.length > WORKBOOK_MAX_CELLS || rows.length >= WORKBOOK_MAX_ROWS)
      throw new Error("This worksheet has reached the row limit.");
    const source =
      sourceRowNumbers[request.index] ?? (sourceRowNumbers.at(-1) ?? sheet.headerRowNumber) + 1;
    worksheet?.spliceRows(source, 0, []);
    if (worksheet) worksheet.getRow(source).getCell(sheet.firstColumn).value = "";
    sourceRowNumbers = sourceRowNumbers.map((number) => (number >= source ? number + 1 : number));
    sourceRowNumbers.splice(request.index, 0, source);
    rows.splice(request.index, 0, Array(columns.length).fill(null));
  } else if (request.type === "deleteRows") {
    if (
      !indexes.length ||
      indexes.some((index) => !Number.isInteger(index) || index < 0 || index >= rows.length)
    )
      throw new Error("Select rows to remove.");
    const sourceToDelete = indexes.map((index) => sourceRowNumbers[index]!);
    for (const source of [...sourceToDelete].sort((a, b) => b - a))
      worksheet?.spliceRows(source, 1);
    const deletedIndexes = new Set(indexes);
    const deletedSources = sourceToDelete.sort((a, b) => a - b);
    let deletedBefore = 0;
    rows = rows.filter((_, index) => !deletedIndexes.has(index));
    sourceRowNumbers = sourceRowNumbers.flatMap((number, index) => {
      if (deletedIndexes.has(index)) return [];
      while ((deletedSources[deletedBefore] ?? Infinity) < number) deletedBefore++;
      return [number - deletedBefore];
    });
  } else if (request.type === "insertColumn") {
    const name = request.name.trim();
    if (!name) throw new Error("Enter a column name.");
    if (columns.some((column) => column.toLocaleLowerCase() === name.toLocaleLowerCase()))
      throw new Error("That column name is already used.");
    if (!Number.isInteger(request.index) || request.index < 0 || request.index > columns.length)
      throw new Error("Choose a valid column position.");
    if (
      columns.length >= WORKBOOK_MAX_COLUMNS ||
      rows.length * (columns.length + 1) > WORKBOOK_MAX_CELLS
    )
      throw new Error("This worksheet has reached the column limit.");
    const physical = sheet.firstColumn + request.index;
    if (worksheet) {
      if (columns.length > 0)
        worksheet.spliceColumns(physical, 0, Array(worksheet.rowCount).fill(null));
      worksheet.getRow(sheet.headerRowNumber).getCell(physical).value = name;
    }
    columns.splice(request.index, 0, name);
    rows = rows.map((row) => {
      row.splice(request.index, 0, null);
      return row;
    });
  } else {
    if (
      !indexes.length ||
      indexes.some((index) => !Number.isInteger(index) || index < 0 || index >= columns.length)
    )
      throw new Error("Select columns to remove.");
    if (indexes.length === columns.length)
      throw new Error("Keep at least one column in this worksheet.");
    for (const index of indexes) worksheet?.spliceColumns(sheet.firstColumn + index, 1);
    const deletedIndexes = new Set(indexes);
    columns = columns.filter((_, index) => !deletedIndexes.has(index));
    rows = rows.map((row) => row.filter((_, index) => !deletedIndexes.has(index)));
  }
  const next: ActiveSheet = {
    ...sheet,
    columns,
    rows,
    sourceRowNumbers,
    overview: analyzeSheet(sheet.name, columns, rows),
  };
  if (workbook) {
    structuralDirty = true;
    sheetOverrides.set(sheet.name, next);
    xlsxTableRanges = new Map([...xlsxTableRanges].filter(([name]) => name !== sheet.name));
  } else csvSheets = csvSheets.map((current) => (current.name === sheet.name ? next : current));
  return next;
}

async function loadWorkbook(
  file: File,
): Promise<{ firstSheet: ActiveSheet; sheetNames: ReadonlyArray<string> }> {
  completeAgentContextCache = null;
  completeAgentContextFailure = null;
  originalXlsx = null;
  xlsxEdits = new Map();
  structuralDirty = false;
  structureWarning = null;
  sheetOverrides = new Map();
  if (file.size > WORKBOOK_MAX_FILE_BYTES) {
    throw new Error("Choose a workbook smaller than 20 MB.");
  }

  const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  if (extension === ".csv" || extension === ".tsv") {
    const text = await file.text();
    const delimiter = extension === ".tsv" ? "\t" : detectCsvDelimiter(text);
    csvDelimiter = delimiter;
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

  originalXlsx = await file.arrayBuffer();
  const normalizedXlsx = await normalizeXlsx(originalXlsx);
  structureWarning = normalizedXlsx.structureWarning;
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

async function normalizeXlsx(source: ArrayBuffer): Promise<{
  data: ArrayBuffer;
  tableRanges: ReadonlyMap<string, XlsxTableRange>;
  structureWarning: string | null;
}> {
  const archive = await JSZip.loadAsync(source);
  const tableRanges = await findWorksheetTableRanges(archive);
  const simpleParts = [
    /^\[Content_Types\]\.xml$/,
    /^_rels\/\.rels$/,
    /^docProps\/(?:app|core)\.xml$/,
    /^xl\/(?:workbook|styles|sharedStrings)\.xml$/,
    /^xl\/_rels\/workbook\.xml\.rels$/,
    /^xl\/theme\/theme\d+\.xml$/,
    /^xl\/worksheets\/sheet\d+\.xml$/,
  ];
  const hasComplexParts = Object.entries(archive.files).some(
    ([name, entry]) => !entry.dir && !simpleParts.some((pattern) => pattern.test(name)),
  );
  let hasFormulas = false;

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
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name) && /<(?:[\w.-]+:)?f(?:\s|>)/.test(xml))
      hasFormulas = true;
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
    structureWarning: hasComplexParts
      ? "This workbook has Excel tables, charts, or other linked parts. Change its rows, columns, or worksheets in Excel to preserve them."
      : hasFormulas
        ? "This workbook has formulas. Change its rows, columns, or worksheets in Excel so formula references stay correct."
        : null,
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

async function saveOriginalXlsx(): Promise<ArrayBuffer> {
  if (!originalXlsx) throw new Error("The original workbook is unavailable.");
  const archive = await JSZip.loadAsync(originalXlsx);
  const workbookXml = await readZipText(archive, "xl/workbook.xml");
  const relationshipsXml = await readZipText(archive, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relationshipsXml) throw new Error("The workbook structure is incomplete.");
  const relationshipTargets = new Map(
    readXmlStartTags(relationshipsXml, "Relationship").flatMap((attributes) => {
      const id = readXmlAttribute(attributes, "Id");
      const target = readXmlAttribute(attributes, "Target");
      return id && target ? [[id, target] as const] : [];
    }),
  );
  for (const attributes of readXmlStartTags(workbookXml, "sheet")) {
    const name = readXmlAttribute(attributes, "name");
    const id = readXmlAttribute(attributes, "r:id");
    const target = id ? relationshipTargets.get(id) : undefined;
    const edits = name ? xlsxEdits.get(decodeXmlText(name)) : undefined;
    if (!edits?.size) continue;
    if (!target) throw new Error(`Could not save worksheet ${name}.`);
    const path = resolvePackagePath("xl/workbook.xml", target);
    const xml = await readZipText(archive, path);
    if (!xml) throw new Error(`Could not save worksheet ${name}.`);
    let updated = xml;
    for (const [address, value] of edits) updated = patchXlsxCell(updated, address, value);
    archive.file(path, updated);
  }
  return archive.generateAsync({ type: "arraybuffer" });
}

function patchXlsxCell(xml: string, address: string, value: WorkbookCell): string {
  const cellPattern = new RegExp(
    `<(?:[\\w.-]+:)?c\\b(?=[^>]*\\br="${address}")[^>]*(?:\\/>|>[\\s\\S]*?<\\/(?:[\\w.-]+:)?c>)`,
  );
  const old = xml.match(cellPattern)?.[0];
  const style = old?.match(/\bs="([^"]+)"/)?.[1];
  const attributes = ` r="${address}"${style ? ` s="${style}"` : ""}`;
  const replacement =
    value === null
      ? `<c${attributes}/>`
      : typeof value === "string"
        ? `<c${attributes} t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
        : `<c${attributes}${typeof value === "boolean" ? ' t="b"' : ""}><v>${typeof value === "boolean" ? Number(value) : value}</v></c>`;
  if (old) return xml.replace(cellPattern, () => replacement);
  const rowNumber = address.match(/\d+$/)?.[0];
  const rowPattern = new RegExp(
    `<(?:[\\w.-]+:)?row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/(?:[\\w.-]+:)?row>`,
  );
  if (!rowNumber || !rowPattern.test(xml)) throw new Error(`Could not locate cell ${address}.`);
  return xml.replace(rowPattern, (row) => row.replace(/<\/(?:[\w.-]+:)?row>$/, `${replacement}$&`));
}

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
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
  const sourceRows: Array<number> = [];
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
      sourceRows.push(rowNumber);
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
      if (width > 0) {
        rawRows.push(values);
        sourceRows.push(row.number);
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
      sourceRows.splice(0, headerCandidate);
      maxColumnCount = rawRows.reduce((largest, row) => Math.max(largest, row.length), 0);
    }
  }
  return makeSheet(
    worksheet.name,
    rawRows,
    maxColumnCount,
    sourceRows,
    tableRange?.firstColumn ?? 1,
  );
}

function parseDelimitedFile(text: string, delimiter: "," | "\t" | ";"): ReadonlyArray<ActiveSheet> {
  const matrix = parseDelimitedText(text, delimiter);
  const rows = matrix;
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
  sourceRows: ReadonlyArray<number> = rows.map((_, index) => index + 1),
  firstColumn = 1,
): ActiveSheet {
  const headerRow = rows[0];
  const columns = makeHeaders(headerRow ?? [], columnCount);
  const tableRows = rows
    .slice(1)
    .map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? null));
  const overview = analyzeSheet(name, columns, tableRows);
  return {
    name,
    columns,
    rows: tableRows,
    overview,
    sourceRowNumbers: sourceRows.slice(1),
    firstColumn,
    headerRowNumber: sourceRows[0] ?? 1,
  };
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

function findRowIndexes(sheet: ActiveSheet, term: string): ReadonlyArray<number> {
  const normalizedTerm = term.trim().toLocaleLowerCase();
  return sheet.rows.flatMap((row, index) =>
    !normalizedTerm ||
    row.some((value) => formatCell(value).toLocaleLowerCase().includes(normalizedTerm))
      ? [index]
      : [],
  );
}

function answerQuestion(sheet: ActiveSheet, question: string): string {
  const normalizedQuestion = normalizeForMatching(question.trim());
  if (!normalizedQuestion) return "Enter a question about this worksheet.";
  const dutch =
    /\b(wat|welke|elke|ieder|iedere|alle|hoeveel|inkomen|gemiddelde|totaal|soort|blad|pagina)\b/.test(
      normalizedQuestion,
    );
  const asksAllSheets =
    /\b(each|every|all|elk|elke|ieder|iedere|alle)\b/.test(normalizedQuestion) &&
    /\b(sheet|worksheet|tab|page|blad|pagina)\b/.test(normalizedQuestion);
  const asksForTypes = /\b(type|types|datatype|datatypes|kind|kinds|soort|soorten|format)\b/.test(
    normalizedQuestion,
  );

  if (asksForTypes) {
    if (asksAllSheets) {
      return formatWorkbookSheets(sheet, (current) => formatColumnTypes(current, dutch));
    }
    const column = findQuestionColumn(sheet.overview.columns, normalizedQuestion);
    if (column) {
      const examples =
        column.examples.length > 0 ? ` Examples: ${column.examples.join(", ")}.` : "";
      return dutch
        ? `${column.name} bevat ${dutchColumnType(column.type).toLocaleLowerCase()}.${examples}`
        : `${column.name} contains ${column.type.toLocaleLowerCase()} data.${examples}`;
    }
    return formatColumnTypes(sheet, dutch);
  }

  if (asksAllSheets) {
    return formatWorkbookSheets(sheet, (current) => formatSheetSummary(current, dutch));
  }

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
    /\b(how many|count|number of|records|customers|rows|aantal|rijen)\b/.test(normalizedQuestion) ||
    (/\bhoeveel\b/.test(normalizedQuestion) &&
      /\b(records|customers|rows|aantal|rijen)\b/.test(normalizedQuestion))
  ) {
    return `This worksheet has ${formatCell(sheet.overview.rowCount)} rows across ${formatCell(sheet.overview.columns.length)} columns.`;
  }

  const column = findQuestionColumn(sheet.overview.columns, normalizedQuestion);
  if (column?.type === "Number" && column.average !== null) {
    if (/\b(average|avg|mean|gemiddelde)\b/.test(normalizedQuestion)) {
      return `The average ${column.name} is ${formatNumber(column.average)}.`;
    }
    if (/\b(total|sum|totaal|som|how much|hoeveel|what is the total)\b/.test(normalizedQuestion)) {
      return dutch
        ? `Het totaal voor ${column.name} is ${formatNumber(column.total ?? 0)}.`
        : `The total ${column.name} is ${formatNumber(column.total ?? 0)}.`;
    }
    if (/\b(minimum|lowest|smallest|min|laagste|kleinste)\b/.test(normalizedQuestion)) {
      return `The lowest ${column.name} is ${formatNumber(column.minimum ?? 0)}.`;
    }
    if (/\b(maximum|highest|largest|max|hoogste|grootste)\b/.test(normalizedQuestion)) {
      return `The highest ${column.name} is ${formatNumber(column.maximum ?? 0)}.`;
    }
  }

  if (/\b(how much|hoeveel)\b/.test(normalizedQuestion)) {
    const numericColumns = sheet.overview.columns.filter(
      (candidate) => candidate.type === "Number" && candidate.total !== null,
    );
    if (numericColumns.length === 1) {
      const onlyColumn = numericColumns[0];
      return dutch
        ? `Het totaal voor ${onlyColumn?.name ?? ""} is ${formatNumber(onlyColumn?.total ?? 0)}.`
        : `The total ${onlyColumn?.name ?? ""} is ${formatNumber(onlyColumn?.total ?? 0)}.`;
    }
    if (numericColumns.length > 1) {
      const fields = numericColumns.map((candidate) => candidate.name).join(", ");
      return dutch
        ? `Welk numeriek veld bedoel je? Beschikbare velden: ${fields}.`
        : `Which numeric field do you mean? Available fields: ${fields}.`;
    }
  }

  return "Try asking what each sheet contains, which data types its columns use, or for counts, missing values, duplicates, and numeric totals or averages.";
}

function formatWorkbookSheets(active: ActiveSheet, format: (sheet: ActiveSheet) => string): string {
  if (!workbook) return csvSheets.map(format).join("\n");

  const summaries: Array<string> = [];
  for (const worksheet of workbook.worksheets) {
    const sheet =
      worksheet.name === active.name
        ? active
        : makeXlsxSheet(worksheet, xlsxTableRanges.get(worksheet.name));
    summaries.push(format(sheet));
  }
  return summaries.join("\n");
}

function normalizeForMatching(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase();
}

function findQuestionColumn(
  columns: ReadonlyArray<WorkbookColumnProfile>,
  normalizedQuestion: string,
): WorkbookColumnProfile | undefined {
  return [...columns]
    .sort((left, right) => right.name.length - left.name.length)
    .find((candidate) => normalizedQuestion.includes(normalizeForMatching(candidate.name)));
}

function formatSheetSummary(sheet: ActiveSheet, dutch: boolean): string {
  const columns = sheet.overview.columns
    .slice(0, 12)
    .map((column) => `${column.name} (${dutch ? dutchColumnType(column.type) : column.type})`)
    .join(", ");
  const moreColumns =
    sheet.overview.columns.length > 12
      ? dutch
        ? ` en ${String(sheet.overview.columns.length - 12)} meer`
        : `, and ${String(sheet.overview.columns.length - 12)} more`
      : "";
  const columnDetails = columns ? `: ${columns}${moreColumns}` : "";
  return dutch
    ? `${sheet.name}: ${formatCell(sheet.overview.rowCount)} rijen, ${formatCell(sheet.overview.columns.length)} kolommen${columnDetails}.`
    : `${sheet.name}: ${formatCell(sheet.overview.rowCount)} rows, ${formatCell(sheet.overview.columns.length)} columns${columnDetails}.`;
}

function formatColumnTypes(sheet: ActiveSheet, dutch: boolean): string {
  const fields = sheet.overview.columns
    .slice(0, 20)
    .map((column) => `${column.name}: ${dutch ? dutchColumnType(column.type) : column.type}`)
    .join(", ");
  const moreColumns =
    sheet.overview.columns.length > 20
      ? dutch
        ? `, en ${String(sheet.overview.columns.length - 20)} meer`
        : `, and ${String(sheet.overview.columns.length - 20)} more`
      : "";
  return dutch
    ? `${sheet.name}, gegevenstypen per kolom: ${fields}${moreColumns}.`
    : `${sheet.name}, data types by column: ${fields}${moreColumns}.`;
}

function dutchColumnType(type: WorkbookColumnProfile["type"]): string {
  switch (type) {
    case "Number":
      return "Getal";
    case "Date":
      return "Datum";
    case "Text":
      return "Tekst";
    case "Mixed":
      return "Gemengd";
  }
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
    structureWarning,
    sheetNames,
    overview: sheet.overview,
    rows: sheet.rows.slice(0, WORKBOOK_PAGE_SIZE),
    rowIndexes: sheet.rows.slice(0, WORKBOOK_PAGE_SIZE).map((_, index) => index),
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

function quoteDelimitedCell(value: WorkbookCell | undefined, delimiter: string): string {
  const text = formatCell(value);
  return text.includes(delimiter) || /["\r\n]/.test(text)
    ? `"${text.replaceAll('"', '""')}"`
    : text;
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
