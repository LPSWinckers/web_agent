import { describe, expect, it } from "vite-plus/test";

import { presentationAgentContext, wordAgentContext } from "./editorAgentContext";

describe("embedded file chat context", () => {
  it("points a presentation turn at its editable source and export", () => {
    const text = presentationAgentContext("powerpoints/Plan.t3deck.json", 2, "chart");
    expect(text).toContain("powerpoints/Plan.t3deck.json");
    expect(text).toContain("powerpoints/Plan.pptx");
    expect(text).toContain("Current slide: 2, selected element: chart");
    expect(text).toContain("regenerates its .pptx export");
  });

  it("points a Word turn at the open document", () => {
    expect(wordAgentContext("reports/Brief.docx", true)).toContain("reports/Brief.docx");
    expect(wordAgentContext("reports/Brief.docx", false)).toContain("read-only text preview");
    expect(wordAgentContext("reports/Brief.docx", true, "Selected words")).toContain(
      "Selected words",
    );
  });
});
