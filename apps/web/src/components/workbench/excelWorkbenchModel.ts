export type WorkbookCell = string | number | boolean | null;

export interface WorkbookColumnProfile {
  readonly name: string;
  readonly type: "Number" | "Date" | "Text" | "Mixed";
  readonly nonEmpty: number;
  readonly missing: number;
  readonly unique: number;
  readonly examples: ReadonlyArray<string>;
  readonly total: number | null;
  readonly average: number | null;
  readonly minimum: number | null;
  readonly maximum: number | null;
}

export interface WorkbookSheetOverview {
  readonly name: string;
  readonly rowCount: number;
  readonly columns: ReadonlyArray<WorkbookColumnProfile>;
  readonly duplicateRows: number;
  readonly missingCells: number;
  readonly completeness: number;
}

export interface WorkbookChartPoint {
  readonly label: string;
  readonly total: number;
  readonly records: number;
}

export type WorkbookCommand =
  | { readonly type: "load"; readonly file: File }
  | { readonly type: "sheet"; readonly name: string }
  | { readonly type: "query"; readonly term: string; readonly page: number }
  | { readonly type: "ask"; readonly question: string }
  | { readonly type: "chart"; readonly groupColumn: string; readonly valueColumn: string }
  | { readonly type: "context" }
  | { readonly type: "agentContext" }
  | {
      readonly type: "editCell";
      readonly row: number;
      readonly column: number;
      readonly value: WorkbookCell;
    }
  | { readonly type: "insertRow"; readonly index: number }
  | { readonly type: "deleteRows"; readonly indexes: ReadonlyArray<number> }
  | { readonly type: "insertColumn"; readonly index: number; readonly name: string }
  | { readonly type: "deleteColumns"; readonly indexes: ReadonlyArray<number> }
  | { readonly type: "addSheet"; readonly name: string }
  | { readonly type: "save" }
  | { readonly type: "export" };

type WithRequestId<Command> = Command extends WorkbookCommand
  ? Command & { readonly requestId: number }
  : never;

export type WorkbookRequest = WithRequestId<WorkbookCommand>;

export type WorkbookResponse =
  | {
      readonly type: "loaded" | "sheet";
      readonly requestId: number;
      readonly fileName: string;
      readonly fileSize: number;
      readonly structureWarning: string | null;
      readonly sheetNames: ReadonlyArray<string>;
      readonly overview: WorkbookSheetOverview;
      readonly rows: ReadonlyArray<ReadonlyArray<WorkbookCell>>;
      readonly rowIndexes: ReadonlyArray<number>;
      readonly page: number;
      readonly filteredRowCount: number;
    }
  | {
      readonly type: "rows";
      readonly requestId: number;
      readonly rows: ReadonlyArray<ReadonlyArray<WorkbookCell>>;
      readonly rowIndexes: ReadonlyArray<number>;
      readonly page: number;
      readonly filteredRowCount: number;
    }
  | { readonly type: "answer"; readonly requestId: number; readonly text: string }
  | { readonly type: "context"; readonly requestId: number; readonly text: string }
  | { readonly type: "saved"; readonly requestId: number; readonly data: ArrayBuffer }
  | {
      readonly type: "chart";
      readonly requestId: number;
      readonly points: ReadonlyArray<WorkbookChartPoint>;
    }
  | {
      readonly type: "exported";
      readonly requestId: number;
      readonly fileName: string;
      readonly data: ArrayBuffer;
    }
  | { readonly type: "error"; readonly requestId: number; readonly message: string };

export const WORKBOOK_PAGE_SIZE = 50;
export const WORKBOOK_MAX_FILE_BYTES = 20 * 1024 * 1024;
export const WORKBOOK_MAX_ROWS = 100_000;
export const WORKBOOK_MAX_COLUMNS = 80;
export const WORKBOOK_MAX_CELLS = 1_500_000;

export function formatWorkbookCell(value: WorkbookCell): string {
  if (value === null) return "";
  if (typeof value === "number") {
    return new Intl.NumberFormat(undefined, { maximumFractionDigits: 3 }).format(value);
  }
  if (typeof value === "boolean") return value ? "Yes" : "No";
  return value;
}

export function formatWorkbookNumber(value: number): string {
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(value);
}
