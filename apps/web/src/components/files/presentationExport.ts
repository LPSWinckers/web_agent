import { CONSULTANCY_PRESENTATION_STANDARD } from "@t3tools/shared/consultancyPresentationStandard";
import type PptxGen from "pptxgenjs";
import type { PresentationDeck, PresentationSlide } from "./presentationDeck";

const WIDTH = 13.333;
const { style, colors } = CONSULTANCY_PRESENTATION_STANDARD;

function text(
  slide: PptxGen.Slide,
  value: string,
  x: number,
  y: number,
  w: number,
  h: number,
  size: number,
  color: string,
  bold = false,
) {
  slide.addText(value, {
    x,
    y,
    w,
    h,
    fontSize: size,
    color,
    bold,
    breakLine: false,
    margin: 0,
    valign: "middle",
    fit: "shrink",
    lineSpacingMultiple: 1.1,
  });
}

function addBody(
  slide: PptxGen.Slide,
  body: string,
  x: number,
  y: number,
  w: number,
  h: number,
  color: string,
) {
  slide.addText(body, {
    x,
    y,
    w,
    h,
    fontFace: style.fontFace,
    fontSize: body.length > 350 ? 17 : 20,
    color,
    margin: 0,
    valign: "top",
    fit: "shrink",
    lineSpacingMultiple: 1.12,
  });
}

function addSlide(pptx: PptxGen, item: PresentationSlide, index: number) {
  const slide = pptx.addSlide();
  const headline = item.layout === "cover" || item.layout === "section";
  slide.background = { color: headline ? colors.navy : style.background };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: headline ? 0.16 : WIDTH,
    h: headline ? 7.5 : 0.11,
    line: { color: headline ? colors.gold : colors.navy },
    fill: { color: headline ? colors.gold : colors.navy },
  });
  text(
    slide,
    "BERENSCHOT 2026  /  PRESENTATIE",
    0.88,
    headline ? 0.72 : 0.45,
    9,
    0.24,
    9,
    headline ? "C7D7EA" : style.muted,
    true,
  );
  if (!headline) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.76,
      y: 0.45,
      w: 0.07,
      h: 0.24,
      line: { color: colors.gold },
      fill: { color: colors.gold },
    });
  }
  const titleY = headline ? 1.78 : 0.92;
  text(
    slide,
    item.title,
    0.86,
    titleY,
    11.55,
    headline ? 1.55 : 0.66,
    headline ? 38 : 28,
    headline ? style.background : style.foreground,
    true,
  );
  if (headline) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.89,
      y: 3.63,
      w: 1.2,
      h: 0.08,
      line: { color: colors.gold },
      fill: { color: colors.gold },
    });
    text(slide, item.body, 0.89, 4.03, 11.35, 1.4, 20, "DCE7F2");
  } else if (item.layout === "two-column") {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.85,
      y: 2.12,
      w: 5.53,
      h: 4.23,
      line: { color: colors.paleBlue },
      fill: { color: colors.paleBlue },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 6.77,
      y: 2.12,
      w: 5.55,
      h: 4.23,
      line: { color: "D8E1EA" },
      fill: { color: style.background },
    });
    text(slide, "INZICHT", 1.15, 2.42, 4.8, 0.26, 10, style.accent, true);
    text(slide, "IMPLICATIE", 7.07, 2.42, 4.8, 0.26, 10, colors.warmBrown, true);
    addBody(slide, item.body, 1.15, 2.9, 4.92, 2.9, style.foreground);
    addBody(slide, item.rightBody ?? "", 7.07, 2.9, 4.95, 2.9, style.foreground);
  } else if (item.layout === "chart" && item.chart?.series.length && item.chart.categories.length) {
    const chart = item.chart;
    const percentFormat =
      chart.unit === "%"
        ? chart.series.some((series) => series.values.some((value) => !Number.isInteger(value)))
          ? '0.0"%"'
          : '0"%"'
        : null;
    const chartType = pptx.ChartType[chart.type];
    const circular = chart.type === "pie" || chart.type === "doughnut";
    slide.addChart(
      chartType,
      chart.series.map((series) => ({
        name: series.name,
        labels: chart.categories,
        values: series.values,
      })),
      {
        x: 0.82,
        y: 2.02,
        w: 11.73,
        h: 4.2,
        barDir: "bar",
        showLegend: !circular && chart.type !== "scatter" && chart.series.length > 1,
        showTitle: false,
        showValue: true,
        dataLabelPosition: chart.type === "scatter" ? "r" : circular ? "bestFit" : "outEnd",
        ...(chart.type === "scatter" ? { lineSize: 0, lineDataSymbol: "circle" } : {}),
        dataLabelColor: style.foreground,
        dataLabelFontFace: style.fontFace,
        dataLabelFontSize: 11,
        ...(percentFormat ? { dataLabelFormatCode: percentFormat } : {}),
        catAxisLabelFontFace: style.fontFace,
        catAxisLabelFontSize: chart.categories.length > 10 ? 9 : 12,
        catAxisLabelColor: style.foreground,
        valAxisLabelFontFace: style.fontFace,
        valAxisLabelFontSize: 10,
        ...(chart.type === "bar" || chart.type === "area" ? { valAxisMinVal: 0 } : {}),
        ...(percentFormat ? { valAxisLabelFormatCode: percentFormat } : {}),
        valGridLine: { color: "D8E1EA", size: 0.6 },
        chartColors: [style.accent, style.foreground, colors.cyan, colors.gold],
        border: { color: style.background, pt: 0 },
      },
    );
    if (item.body) text(slide, item.body, 0.89, 6.32, 11.4, 0.37, 12, style.foreground);
  } else {
    addBody(slide, item.body, 0.89, 2.12, 11.55, 4.4, style.foreground);
  }
  if (!headline) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.86,
      y: 1.73,
      w: 0.72,
      h: 0.055,
      line: { color: colors.gold },
      fill: { color: colors.gold },
    });
  }
  const source = item.source ?? item.chart?.source;
  if (source)
    text(slide, `Bron: ${source}`, 0.87, 6.97, 10.65, 0.2, 8, headline ? "C7D7EA" : style.muted);
  text(slide, `${index + 1}`, 12.05, 7.01, 0.45, 0.2, 9, headline ? "C7D7EA" : style.muted);
  slide.addNotes(source ? `Bron: ${source}` : item.body);
}

export async function exportPresentation(deck: PresentationDeck): Promise<string> {
  const { default: pptxgen } = await import("pptxgenjs");
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "T3 Code";
  pptx.subject = deck.title;
  pptx.title = deck.title;
  pptx.theme = { headFontFace: style.fontFace, bodyFontFace: style.fontFace };
  deck.slides.forEach((slide, index) => addSlide(pptx, slide, index));
  const result = await pptx.write({ outputType: "base64", compression: true });
  if (typeof result !== "string")
    throw new Error("PowerPoint export returned an unexpected result.");
  return result;
}

export function downloadPresentation(base64: string, name: string): void {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  const url = URL.createObjectURL(
    new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }),
  );
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
