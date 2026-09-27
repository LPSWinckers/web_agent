export const CONSULTANCY_CHART_TYPES = [
  { type: "bar", name: "Staafgrafiek", description: "Vergelijk categorieën of locaties." },
  { type: "line", name: "Lijngrafiek", description: "Toon ontwikkeling door de tijd." },
  { type: "area", name: "Vlakgrafiek", description: "Toon de omvang van een trend." },
  { type: "pie", name: "Cirkeldiagram", description: "Toon delen van één totaal." },
  {
    type: "doughnut",
    name: "Ringdiagram",
    description: "Toon een verdeling met nadruk op het totaal.",
  },
  {
    type: "scatter",
    name: "Spreidingsdiagram",
    description: "Toon het verband tussen twee getallen.",
  },
] as const;

export type ConsultancyChartType = (typeof CONSULTANCY_CHART_TYPES)[number]["type"];

export interface ConsultancyChart {
  type: ConsultancyChartType;
  categories: string[];
  series: Array<{ name: string; values: number[] }>;
  title?: string;
  unit?: string;
  source?: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Parse the same bounded chart data in chat and presentation sources. */
export function parseConsultancyChart(value: unknown): ConsultancyChart | null {
  if (!isRecord(value) || !CONSULTANCY_CHART_TYPES.some((item) => item.type === value.type))
    return null;
  const categories = value.categories;
  const series = value.series;
  if (
    !Array.isArray(categories) ||
    categories.length < 1 ||
    categories.length > 100 ||
    !categories.every((item) => typeof item === "string" && item.length <= 200) ||
    !Array.isArray(series) ||
    series.length < 1 ||
    series.length > 4 ||
    !series.every(
      (item) =>
        isRecord(item) &&
        typeof item.name === "string" &&
        item.name.length <= 120 &&
        Array.isArray(item.values) &&
        item.values.length === categories.length &&
        item.values.every((number) => typeof number === "number" && Number.isFinite(number)),
    ) ||
    (value.type === "scatter" && (series.length !== 2 || categories.length < 2)) ||
    ((value.type === "pie" || value.type === "doughnut") &&
      (series.length !== 1 || series[0].values.some((number: number) => number < 0)))
  )
    return null;
  for (const key of ["title", "unit", "source"] as const) {
    if (value[key] !== undefined && (typeof value[key] !== "string" || value[key].length > 500))
      return null;
  }
  return {
    type: value.type as ConsultancyChartType,
    categories,
    series: series as ConsultancyChart["series"],
    ...(typeof value.title === "string" ? { title: value.title } : {}),
    ...(typeof value.unit === "string" ? { unit: value.unit } : {}),
    ...(typeof value.source === "string" ? { source: value.source } : {}),
  };
}

export function parseConsultancyChartJson(source: string): ConsultancyChart | null {
  if (source.length > 60_000) return null;
  try {
    return parseConsultancyChart(JSON.parse(source));
  } catch {
    return null;
  }
}

/** Keep ordinary Markdown around complete chart fences intact on native clients. */
export function splitConsultancyChartMarkdown(
  markdown: string,
): Array<
  | { kind: "markdown"; markdown: string; sourceOffset: number }
  | { kind: "chart"; chart: ConsultancyChart; sourceOffset: number }
> {
  const segments: Array<
    | { kind: "markdown"; markdown: string; sourceOffset: number }
    | { kind: "chart"; chart: ConsultancyChart; sourceOffset: number }
  > = [];
  const fence = /(?:^|\n)```t3-chart[ \t]*\r?\n([\s\S]*?)\r?\n```(?=\n|$)/g;
  let start = 0;
  for (const match of markdown.matchAll(fence)) {
    const chart = parseConsultancyChartJson(match[1]!);
    if (!chart || match.index === undefined) continue;
    const chartStart = match.index + (match[0].startsWith("\n") ? 1 : 0);
    if (chartStart > start)
      segments.push({
        kind: "markdown",
        markdown: markdown.slice(start, chartStart),
        sourceOffset: start,
      });
    segments.push({ kind: "chart", chart, sourceOffset: chartStart });
    start = match.index + match[0].length;
  }
  if (start < markdown.length)
    segments.push({ kind: "markdown", markdown: markdown.slice(start), sourceOffset: start });
  return segments;
}
