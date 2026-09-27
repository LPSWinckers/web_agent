import { describe, expect, it } from "vite-plus/test";
import { importChart } from "./presentationChartImport";

describe("presentation chart import", () => {
  it("keeps labeled values and series names from CSV data", async () => {
    const file = new File(["Year,Base,Recommended\n2026,20,25\n2027,24,30\n"], "revenue.csv", {
      type: "text/csv",
    });
    expect(await importChart(file)).toEqual({
      type: "bar",
      categories: ["2026", "2027"],
      series: [
        { name: "Base", values: [20, 24] },
        { name: "Recommended", values: [25, 30] },
      ],
      source: "revenue.csv",
    });
  });
});
