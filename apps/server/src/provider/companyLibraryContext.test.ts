import { describe, expect, it } from "vite-plus/test";
import { DEFAULT_COMPANY_LIBRARY } from "@t3tools/contracts";

import { companyLibraryContext } from "./companyLibraryContext.ts";

const library = {
  ...DEFAULT_COMPANY_LIBRARY,
  companyName: "Acme",
  guidance: "Use our approved colors.",
  chartTemplates: [
    { id: "trend", name: "Monthly trend", instructions: "Use a line chart for monthly values." },
  ],
  skills: [{ id: "review", name: "Review brief", instructions: "Check the source numbers first." }],
  powerpointStandards: [
    { id: "deck", name: "Board deck", instructions: "Use a single takeaway per slide." },
  ],
};

describe("company library agent context", () => {
  it("keeps an empty server library out of turns", () => {
    expect(companyLibraryContext(DEFAULT_COMPANY_LIBRARY, "Fix this test")).toBeNull();
  });

  it("keeps promptless continuation turns promptless", () => {
    expect(companyLibraryContext(library, undefined)).toBeNull();
  });

  it("uses the shared standard for presentation requests despite old custom records", () => {
    const context = companyLibraryContext(library, "Make an Excel graph for a PowerPoint deck");
    expect(context).toContain("Single shared standard: Berenschot 2026");
    expect(context).toContain("background: 'FFFFFF'");
    expect(context).toContain("Slide components:");
    expect(context).toContain("Chart components:");
    expect(context).toContain("Use a line chart for monthly values.");
    expect(context).toContain("Use a single takeaway per slide.");
    expect(context).not.toContain("Check the source numbers first.");
  });

  it("applies the standard with empty settings and Dutch presentation requests", () => {
    const context = companyLibraryContext(null, "Maak een presentatie met een dia per werkblad");
    expect(context).toContain("Berenschot 2026");
    expect(context).toContain("Slide components:");
    expect(context).toContain("Chart components:");
  });

  it("includes the seven sourced graphs for a report presentation", () => {
    const context = companyLibraryContext(
      DEFAULT_COMPANY_LIBRARY,
      "Maak een PowerPoint over het Berenschot Trendsonderzoek",
    );
    expect(context).toContain("Waardepropositie uitgewerkt 41%");
    expect(context).toContain("Financiële dienstverlening 100%");
    expect(context).toContain("Zuid-Europa 19%");
  });

  it("recognizes a named graph without a chart keyword", () => {
    const context = companyLibraryContext(DEFAULT_COMPANY_LIBRARY, "Voeg AI-toepassingen toe");
    expect(context).toContain("Tekst samenvatten 60%");
    expect(context).not.toContain("Zuid-Europa 19%");
  });

  it("includes a skill when the request names it", () => {
    const context = companyLibraryContext(library, "Use Review brief for this report");
    expect(context).toContain("Check the source numbers first.");
    expect(context).not.toContain("Single shared standard:");
  });

  it("provides the saved Word theme and structure for document requests", () => {
    const context = companyLibraryContext(
      {
        ...library,
        wordStandard: {
          ...library.wordStandard,
          fontFamily: "Aptos",
          sections: ["Scope", "Recommendations"],
        },
      },
      "Edit this DOCX report",
    );
    expect(context).toContain("Font: Aptos");
    expect(context).toContain("Scope; Recommendations");
  });

  it("recognizes Dutch report requests", () => {
    expect(companyLibraryContext(null, "Schrijf een rapport voor de klant")).toContain(
      "Word document standard:",
    );
  });
});
