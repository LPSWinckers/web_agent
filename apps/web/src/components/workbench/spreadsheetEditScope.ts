export type EditTarget = "workbook" | "sheet" | "cells" | "rows" | "columns";

export interface SpreadsheetEditScope {
  readonly target: EditTarget;
  readonly sheet: string;
  readonly cells: ReadonlySet<string>;
  readonly rows: ReadonlySet<number>;
  readonly columns: ReadonlySet<number>;
}

export function cellKey(sheet: string, row: number, column: number): string {
  return JSON.stringify([sheet, row, column]);
}

export function canEditCell(
  scope: SpreadsheetEditScope,
  sheet: string,
  row: number,
  column: number,
): boolean {
  const selected =
    scope.target === "workbook" ||
    (scope.target === "sheet" && scope.sheet === sheet) ||
    (scope.target === "cells" && scope.cells.has(cellKey(sheet, row, column))) ||
    (scope.target === "rows" && scope.sheet === sheet && scope.rows.has(row)) ||
    (scope.target === "columns" && scope.sheet === sheet && scope.columns.has(column));
  return selected;
}

export function describeEditScope(
  scope: SpreadsheetEditScope,
  columns: ReadonlyArray<string>,
): string {
  const target =
    scope.target === "workbook"
      ? "all worksheets"
      : scope.target === "sheet"
        ? `worksheet ${JSON.stringify(scope.sheet)}`
        : scope.target === "cells"
          ? `only these cells (worksheet, data row, field): ${[...scope.cells]
              .map((key) => {
                const [sheet, row, column] = JSON.parse(key) as [string, number, number];
                return `${JSON.stringify(sheet)} row ${row + 1} ${JSON.stringify(columns[column] ?? `Column ${column + 1}`)}`;
              })
              .join(", ")}`
          : scope.target === "rows"
            ? `worksheet ${JSON.stringify(scope.sheet)}, data rows ${[...scope.rows].map((row) => row + 1).join(", ")}`
            : `worksheet ${JSON.stringify(scope.sheet)}, columns ${[...scope.columns].map((column) => JSON.stringify(columns[column] ?? column + 1)).join(", ")}`;
  return target;
}
