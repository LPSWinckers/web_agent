import { describe, expect, it } from "vite-plus/test";

import {
  canEditCell,
  cellKey,
  describeEditScope,
  type SpreadsheetEditScope,
} from "./spreadsheetEditScope";

const scope: SpreadsheetEditScope = {
  target: "cells",
  sheet: "Budget",
  cells: new Set([cellKey("Budget", 1, 2)]),
  rows: new Set(),
  columns: new Set(),
};

describe("spreadsheet edit scope", () => {
  it("allows only selected cells", () => {
    expect(canEditCell(scope, "Budget", 1, 2)).toBe(true);
    expect(canEditCell(scope, "Budget", 2, 2)).toBe(false);
    expect(canEditCell(scope, "Other", 1, 2)).toBe(false);
  });

  it("describes the boundary sent to the agent", () => {
    expect(describeEditScope(scope, ["Year", "Name", "Amount"])).toContain(
      '"Budget" row 2 "Amount"',
    );
  });

  it("includes every cell in a selected row or column", () => {
    const rowScope: SpreadsheetEditScope = { ...scope, target: "rows", rows: new Set([1]) };
    expect(canEditCell(rowScope, "Budget", 1, 0)).toBe(true);
    expect(canEditCell(rowScope, "Budget", 1, 5)).toBe(true);
    expect(canEditCell(rowScope, "Budget", 0, 0)).toBe(false);
    const columnScope: SpreadsheetEditScope = {
      ...scope,
      target: "columns",
      columns: new Set([2]),
    };
    expect(canEditCell(columnScope, "Budget", 0, 2)).toBe(true);
    expect(canEditCell(columnScope, "Budget", 10, 2)).toBe(true);
    expect(canEditCell(columnScope, "Budget", 0, 1)).toBe(false);
  });
});
