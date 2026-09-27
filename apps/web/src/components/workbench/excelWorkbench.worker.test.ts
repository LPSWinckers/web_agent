import ExcelJS from "exceljs";
import JSZip from "jszip";
import { expect, it, vi } from "vite-plus/test";

import type { WorkbookRequest, WorkbookResponse } from "./excelWorkbenchModel";

it("edits a data cell and saves the workbook with other sheets intact", async () => {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet("Budget").addRows([
    ["Year", "Amount", "Note"],
    [2024, 10],
    [2025, 20],
  ]);
  workbook.addWorksheet("Notes").addRows([["Name"], ["Keep me"]]);
  const simpleBytes = await workbook.xlsx.writeBuffer();
  const archive = await JSZip.loadAsync(simpleBytes);
  archive.file("xl/custom-preserve.xml", "<preserved/>");
  const source = new File([await archive.generateAsync({ type: "arraybuffer" })], "budget.xlsx");
  let onmessage: ((event: MessageEvent<WorkbookRequest>) => void) | null = null;
  let receive: ((response: WorkbookResponse) => void) | null = null;
  vi.stubGlobal("self", {
    get onmessage() {
      return onmessage;
    },
    set onmessage(value) {
      onmessage = value;
    },
    postMessage(response: WorkbookResponse) {
      receive?.(response);
    },
  });
  try {
    await import("./excelWorkbench.worker");
    const send = (request: WorkbookRequest) =>
      new Promise<WorkbookResponse>((resolve) => {
        receive = resolve;
        onmessage?.({ data: request } as MessageEvent<WorkbookRequest>);
      });
    expect((await send({ type: "load", file: source, requestId: 1 })).type).toBe("loaded");
    const filtered = await send({ type: "query", term: "2025", page: 0, requestId: 4 });
    expect(filtered.type === "rows" ? filtered.rowIndexes : null).toEqual([1]);
    const edit = await send({ type: "editCell", row: 1, column: 1, value: 25, requestId: 2 });
    expect(edit.type).toBe("sheet");
    expect(
      (await send({ type: "editCell", row: 1, column: 2, value: "A & B", requestId: 5 })).type,
    ).toBe("sheet");
    const saved = await send({ type: "save", requestId: 3 });
    expect(saved.type).toBe("saved");
    if (saved.type !== "saved") return;
    const output = new ExcelJS.Workbook();
    await output.xlsx.load(saved.data);
    expect(output.getWorksheet("Budget")?.getCell("B3").value).toBe(25);
    expect(output.getWorksheet("Budget")?.getCell("C3").value).toBe("A & B");
    expect(output.getWorksheet("Notes")?.getCell("A2").value).toBe("Keep me");
    expect(
      await (await JSZip.loadAsync(saved.data)).file("xl/custom-preserve.xml")?.async("string"),
    ).toBe("<preserved/>");
    expect((await send({ type: "insertRow", index: 1, requestId: 24 })).type).toBe("error");
    expect(
      (
        await send({
          type: "load",
          file: new File([new Uint8Array(simpleBytes)], "budget.xlsx"),
          requestId: 25,
        })
      ).type,
    ).toBe("loaded");
    expect(
      (await send({ type: "editCell", row: 1, column: 1, value: 25, requestId: 26 })).type,
    ).toBe("sheet");

    const insertedRow = await send({ type: "insertRow", index: 1, requestId: 6 });
    expect(insertedRow.type === "sheet" ? insertedRow.overview.rowCount : null).toBe(3);
    expect(
      (await send({ type: "editCell", row: 1, column: 1, value: 15, requestId: 7 })).type,
    ).toBe("sheet");
    const insertedColumn = await send({
      type: "insertColumn",
      index: 2,
      name: "New",
      requestId: 8,
    });
    expect(
      insertedColumn.type === "sheet"
        ? insertedColumn.overview.columns.map((column) => column.name)
        : null,
    ).toEqual(["Year", "Amount", "New", "Note"]);
    expect((await send({ type: "deleteRows", indexes: [0], requestId: 9 })).type).toBe("sheet");
    expect((await send({ type: "deleteColumns", indexes: [3], requestId: 10 })).type).toBe("sheet");
    const addedSheet = await send({ type: "addSheet", name: "Forecast", requestId: 11 });
    expect(addedSheet.type === "sheet" ? addedSheet.sheetNames : null).toEqual([
      "Budget",
      "Notes",
      "Forecast",
    ]);
    const firstColumn = await send({
      type: "insertColumn",
      index: 0,
      name: "Month",
      requestId: 12,
    });
    if (firstColumn.type === "error") throw new Error(firstColumn.message);
    expect(firstColumn.type).toBe("sheet");
    expect((await send({ type: "insertRow", index: 0, requestId: 13 })).type).toBe("sheet");
    const changed = await send({ type: "save", requestId: 14 });
    expect(changed.type).toBe("saved");
    if (changed.type !== "saved") return;
    const reopened = await send({
      type: "load",
      file: new File([changed.data], "budget.xlsx"),
      requestId: 15,
    });
    expect(reopened.type === "loaded" ? reopened.overview.rowCount : null).toBe(2);
    expect(reopened.type === "loaded" ? reopened.rows.map((row) => row[1]) : null).toEqual([
      15, 25,
    ]);
    const forecast = await send({ type: "sheet", name: "Forecast", requestId: 16 });
    expect(
      forecast.type === "sheet" ? forecast.overview.columns.map((column) => column.name) : null,
    ).toEqual(["Month"]);
    expect(forecast.type === "sheet" ? forecast.overview.rowCount : null).toBe(1);

    expect(
      (
        await send({
          type: "load",
          file: new File(["Year,Amount\n2024,10\n2025,20"], "data.csv"),
          requestId: 17,
        })
      ).type,
    ).toBe("loaded");
    expect((await send({ type: "insertRow", index: 1, requestId: 18 })).type).toBe("sheet");
    expect((await send({ type: "insertColumn", index: 1, name: "Note", requestId: 19 })).type).toBe(
      "sheet",
    );
    const csv = await send({ type: "save", requestId: 20 });
    expect(csv.type).toBe("saved");
    if (csv.type !== "saved") return;
    const csvText = new TextDecoder().decode(csv.data);
    expect(csvText).toContain("2024,,10\r\n,,\r\n2025,,20");
    const reopenedCsv = await send({
      type: "load",
      file: new File([csv.data], "data.csv"),
      requestId: 21,
    });
    expect(reopenedCsv.type === "loaded" ? reopenedCsv.overview.rowCount : null).toBe(3);

    archive.file("xl/tables/table1.xml", "<table/>");
    const tableFile = new File(
      [await archive.generateAsync({ type: "arraybuffer" })],
      "table.xlsx",
    );
    const tableLoad = await send({ type: "load", file: tableFile, requestId: 22 });
    expect(tableLoad.type === "loaded" ? tableLoad.structureWarning : null).toContain(
      "Excel tables",
    );
    expect((await send({ type: "insertRow", index: 0, requestId: 23 })).type).toBe("error");
  } finally {
    vi.unstubAllGlobals();
  }
});
