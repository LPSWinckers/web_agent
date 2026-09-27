import { describe, expect, it } from "vite-plus/test";
import JSZip from "jszip";
import { newDeck, newSlide, parseDeck } from "./presentationDeck";
import { exportPresentation } from "./presentationExport";

describe("consultant presentations", () => {
  it("opens presentation sources saved by the earlier editor", () => {
    const deck = parseDeck(
      JSON.stringify({
        version: 1,
        title: "Prior deck",
        theme: "Corporate",
        slides: [{ title: "Summary", body: "Finding" }],
      }),
    );
    expect(deck.style).toEqual({
      name: "Berenschot 2026",
      background: "FFFFFF",
      foreground: "1F345E",
      accent: "0075AB",
      muted: "66758A",
      fontFace: "Arial",
    });
    expect(deck.slides[0]?.title).toBe("Summary");
  });

  it("exports the shared styling, editable text, and a native chart", async () => {
    const deck = newDeck();
    deck.style = {
      name: "Old custom style",
      background: "FF0000",
      foreground: "FF0000",
      accent: "FF0000",
      muted: "FF0000",
      fontFace: "Aptos",
    };
    deck.slides.push({
      ...newSlide("chart"),
      title: "Revenue grows in the recommended case",
      body: "The recommended scenario reaches 30 in 2027.",
      chart: {
        type: "bar",
        categories: ["2026", "2027"],
        series: [{ name: "Revenue", values: [20, 30] }],
        unit: "%",
      },
      source: "revenue.xlsx",
    });

    const base64 = await exportPresentation(deck);
    const zip = await JSZip.loadAsync(base64, { base64: true });
    const slides = Object.keys(zip.files).filter((path) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(path),
    );
    const chartFiles = Object.keys(zip.files).filter((path) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(path),
    );

    expect(slides).toHaveLength(3);
    expect(chartFiles).toHaveLength(1);
    expect(await zip.file(slides[2]!)?.async("string")).toContain(
      "Revenue grows in the recommended case",
    );
    const coverXml = await zip.file(slides[0]!)?.async("string");
    expect(coverXml).toContain('val="1F345E"');
    expect(coverXml).toContain('val="FAB900"');
    expect(coverXml).not.toContain('val="FF0000"');
    const chartXml = await zip.file(chartFiles[0]!)?.async("string");
    expect(chartXml).toContain("Revenue");
    expect(chartXml).toContain('val="0075AB"');
    expect(chartXml).toContain("%");
    expect(parseDeck(JSON.stringify(deck)).style.name).toBe("Berenschot 2026");
    expect(parseDeck(JSON.stringify(deck)).slides[2]?.chart?.unit).toBe("%");
    expect(parseDeck(JSON.stringify(deck)).slides[2]?.chart?.series[0]?.values).toEqual([20, 30]);
  });
});
