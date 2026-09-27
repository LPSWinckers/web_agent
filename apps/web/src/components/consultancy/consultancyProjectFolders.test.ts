import { describe, expect, it } from "vite-plus/test";

import { consultancyProjectFolders } from "./consultancyProjectFolders";

describe("consultancy project setup files", () => {
  it("creates the project folders without adding office documents", () => {
    const files = consultancyProjectFolders();

    expect(files.map((file) => file.path)).toContain("excel/.keep");
    expect(files.map((file) => file.path)).not.toContain("excel/Analyse.xlsx");
    expect(files.some((file) => /\.(docx|pptx)$/i.test(file.path))).toBe(false);
  });
});
