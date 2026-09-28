import type { ThreadId } from "@t3tools/contracts";
import { CONSULTANCY_PRESENTATION_STANDARD } from "@t3tools/shared/consultancyPresentationStandard";
import { parseConsultancyChart, type ConsultancyChart } from "@t3tools/shared/consultancyChart";

export interface PresentationStyle {
  name: string;
  background: string;
  foreground: string;
  accent: string;
  muted: string;
  fontFace: string;
}

export const PRESENTATION_STYLE: PresentationStyle = CONSULTANCY_PRESENTATION_STANDARD.style;

export type SlideLayout = "cover" | "section" | "content" | "two-column" | "chart";
export type PresentationChart = ConsultancyChart;
export interface PresentationImage {
  path: string;
  alt: string;
  sourceItemId?: string;
  sourceUrl?: string;
  keywords?: string[];
  restrictions?: string[];
}
export interface PresentationSlide {
  id: string;
  layout: SlideLayout;
  title: string;
  body: string;
  rightBody?: string;
  chart?: PresentationChart;
  image?: PresentationImage;
  source?: string;
}
export interface PresentationDeck {
  version: 1;
  title: string;
  style: PresentationStyle;
  slides: PresentationSlide[];
  chatThreadId?: ThreadId;
}

export function newSlide(layout: SlideLayout = "content"): PresentationSlide {
  return {
    // @effect-diagnostics-next-line cryptoRandomUUID:off
    id: crypto.randomUUID(),
    layout,
    title: "New slide",
    body: "",
    ...(layout === "chart"
      ? {
          chart: {
            type: "bar" as const,
            categories: ["Categorie A", "Categorie B"],
            series: [{ name: "Waarde", values: [0, 0] }],
          },
        }
      : {}),
  };
}

export function newDeck(): PresentationDeck {
  return {
    version: 1,
    title: "Nieuwe presentatie",
    style: { ...PRESENTATION_STYLE },
    slides: [
      { ...newSlide("cover"), title: "Nieuwe presentatie", body: "[Korte ondertitel]" },
      {
        ...newSlide(),
        title: "Kernboodschap",
        body: "[Inzicht]\n[Implicatie]\n[Volgende stap]",
      },
    ],
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function parseDeck(source: string): PresentationDeck {
  const value: unknown = JSON.parse(source);
  if (!isRecord(value) || value.version !== 1 || typeof value.title !== "string")
    throw new Error("This is not a supported presentation source.");
  if (!value.style && typeof value.theme === "string" && Array.isArray(value.slides)) {
    const slides = value.slides.map((slide, index) => {
      if (!isRecord(slide) || typeof slide.title !== "string" || typeof slide.body !== "string")
        throw new Error(`Slide ${index + 1} is incomplete.`);
      return {
        id: `legacy-${index + 1}`,
        layout: index === 0 ? ("cover" as const) : ("content" as const),
        title: slide.title,
        body: slide.body,
      };
    });
    return { version: 1, title: value.title, style: { ...PRESENTATION_STYLE }, slides };
  }
  if (!Array.isArray(value.slides) || value.slides.length > 200)
    throw new Error("The presentation needs 0 to 200 slides.");
  const slides = value.slides.map((item, index): PresentationSlide => {
    if (
      !isRecord(item) ||
      typeof item.title !== "string" ||
      typeof item.body !== "string" ||
      !["cover", "section", "content", "two-column", "chart"].includes(String(item.layout))
    )
      throw new Error(`Slide ${index + 1} is incomplete.`);
    let parsedChart: PresentationChart | undefined;
    if (item.layout === "chart") {
      const chart = parseConsultancyChart(item.chart);
      if (!chart) throw new Error(`Slide ${index + 1} has invalid chart data.`);
      parsedChart = chart;
    }
    let parsedImage: PresentationImage | undefined;
    if (item.image !== undefined) {
      if (
        item.layout !== "content" ||
        !isRecord(item.image) ||
        typeof item.image.path !== "string" ||
        item.image.path.trim().length === 0 ||
        typeof item.image.alt !== "string"
      ) {
        throw new Error(
          `Slide ${index + 1} has an invalid image. Images are supported on content slides.`,
        );
      }
      parsedImage = {
        path: item.image.path,
        alt: item.image.alt,
        ...(typeof item.image.sourceItemId === "string"
          ? { sourceItemId: item.image.sourceItemId }
          : {}),
        ...(typeof item.image.sourceUrl === "string" ? { sourceUrl: item.image.sourceUrl } : {}),
        ...(Array.isArray(item.image.keywords) &&
        item.image.keywords.every((value) => typeof value === "string")
          ? { keywords: item.image.keywords }
          : {}),
        ...(Array.isArray(item.image.restrictions) &&
        item.image.restrictions.every((value) => typeof value === "string")
          ? { restrictions: item.image.restrictions }
          : {}),
      };
    }
    return {
      id: typeof item.id === "string" ? item.id : `slide-${index + 1}`,
      layout: item.layout as SlideLayout,
      title: item.title,
      body: item.body,
      ...(typeof item.rightBody === "string" ? { rightBody: item.rightBody } : {}),
      ...(typeof item.source === "string" ? { source: item.source } : {}),
      ...(parsedChart ? { chart: parsedChart } : {}),
      ...(parsedImage ? { image: parsedImage } : {}),
    };
  });
  return {
    version: 1,
    title: value.title,
    style: { ...PRESENTATION_STYLE },
    slides,
    ...(typeof value.chatThreadId === "string"
      ? { chatThreadId: value.chatThreadId as ThreadId }
      : {}),
  };
}

export function deckPath(title: string): string {
  const name =
    [...title]
      .filter((character) => character.charCodeAt(0) > 31)
      .join("")
      .replace(/[<>:"/\\|?*]/g, "")
      .trim()
      .replace(/[. ]+$/, "")
      .slice(0, 90) || "Presentation";
  return `powerpoints/${name}`;
}
