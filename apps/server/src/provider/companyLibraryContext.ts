import {
  DEFAULT_COMPANY_LIBRARY,
  type CompanyLibrary,
  type CompanyLibraryEntry,
} from "@t3tools/contracts";
import {
  CONSULTANCY_CHART_COMPONENTS,
  CONSULTANCY_PRESENTATION_GUIDANCE,
  CONSULTANCY_PRESENTATION_STANDARD,
  CONSULTANCY_SLIDE_COMPONENTS,
} from "@t3tools/shared/consultancyPresentationStandard";
import { CONSULTANCY_CHART_TYPES } from "@t3tools/shared/consultancyChart";

const MAX_CONTEXT_CHARS = 6_000;

function matchingEntries(entries: ReadonlyArray<CompanyLibraryEntry>, text: string) {
  return entries.filter((entry) => text.includes(entry.name.toLowerCase()));
}

/** Include relevant records without sending the whole library on every turn. */
export function companyLibraryContext(
  library: CompanyLibrary | null,
  userText: string | undefined,
  hasAttachments = false,
): string | null {
  if (!userText && !hasAttachments) return null;
  const saved = library ?? DEFAULT_COMPANY_LIBRARY;
  const text = userText?.toLowerCase() ?? "";
  const namedCharts = CONSULTANCY_CHART_COMPONENTS.filter((chart) =>
    text.includes(chart.name.toLowerCase()),
  );
  const presentationRequested =
    /\b(powerpoints?|pptx?|slides?|decks?|presentations?|presentaties?|dias?|t3deck)\b/i.test(text);
  const wordRequested =
    /\b(docx?|word|document(?:en|s)?|rapport(?:en)?|verslag(?:en)?|reports?)\b/i.test(text);
  const chartRequested =
    /\b(charts?|graphs?|graphiek(?:en)?|plots?|visuali[sz](?:e|ation|ations)|grafiek(?:en)?|diagram(?:men)?)\b|(?:staaf|lijn|cirkel|ring|vlak|spreidings)(?:grafiek|graphiek)/i.test(
      text,
    ) ||
    (presentationRequested && /\b(data|excel|spreadsheet|werkblad|werkbladen)\b/i.test(text)) ||
    namedCharts.length > 0;
  const reportRequested = /\b(berenschot|trendsonderzoek)\b/i.test(text);
  const skills = matchingEntries(saved.skills, text);
  if (
    !presentationRequested &&
    !wordRequested &&
    !chartRequested &&
    !saved.companyName &&
    !saved.guidance &&
    skills.length === 0
  ) {
    return null;
  }

  const parts = ["Company guidance for this request:"];
  if (saved.companyName) parts.push(`Company: ${saved.companyName}`);
  if (saved.guidance) parts.push(`General guidance:\n${saved.guidance}`);
  if (wordRequested) {
    const standard = saved.wordStandard;
    parts.push(
      `Word document standard: ${standard.name}. Font: ${standard.fontFamily}. Body color: #${standard.bodyColor}; heading color: #${standard.headingColor}; accent color: #${standard.accentColor}. Preferred section order: ${standard.sections.join("; ")}. Preserve the document's existing content and add missing sections only when relevant.`,
    );
  }
  if (presentationRequested || chartRequested) {
    parts.push(
      `Single shared standard: ${CONSULTANCY_PRESENTATION_STANDARD.name}.\n${CONSULTANCY_PRESENTATION_GUIDANCE}`,
    );
    if (presentationRequested)
      parts.push(
        `Slide components: ${CONSULTANCY_SLIDE_COMPONENTS.map((item) => `${item.id} (${item.description})`).join("; ")}.`,
      );
    if (presentationRequested) {
      parts.push(
        "For a presentation open in T3 Code, edit its existing .t3deck.json source so the editor can refresh the PowerPoint export. Search the project workspace for supplied PowerPoint templates and source data before creating new material. Do not claim the open presentation is inaccessible without checking its source path.",
      );
      for (const entry of saved.powerpointStandards) {
        const next = `PowerPoint template "${entry.name}":\n${entry.instructions}`;
        if (parts.join("\n\n").length + next.length + 2 > MAX_CONTEXT_CHARS) break;
        parts.push(next);
      }
    }
    if (chartRequested) {
      for (const entry of saved.chartTemplates) {
        const next = `Chart template "${entry.name}":\n${entry.instructions}`;
        if (parts.join("\n\n").length + next.length + 2 > MAX_CONTEXT_CHARS) break;
        parts.push(next);
      }
    }
    if (chartRequested || reportRequested)
      parts.push(
        `Chart components: ${CONSULTANCY_CHART_TYPES.map((item) => `${item.name} [${item.type}: ${item.description}]`).join("; ")}. One reusable bar component; adapt labels, series, values, unit, and source to the user's data.`,
      );
    if (chartRequested)
      parts.push(
        'To show a chart directly in chat, include a fenced t3-chart JSON block with {"type":"bar","title":"...","categories":["A","B"],"series":[{"name":"Reeks","values":[1,2]}],"unit":"%","source":"..."}. Supported types: bar, line, area, pie, doughnut, scatter. For scatter, use exactly two numeric series: x values followed by y values, with category labels naming the points. Pie and doughnut use exactly one nonnegative series. Ground every value in the supplied files; do not output [visualize], a local HTML link, or raw follow-up directives in place of the chart.',
      );
    if (reportRequested || namedCharts.length)
      for (const chart of reportRequested ? CONSULTANCY_CHART_COMPONENTS : namedCharts) {
        const next = `${chart.name}: ${chart.categories.map((category, index) => `${category} ${chart.values[index]}${chart.unit}`).join("; ")}. Source: ${chart.source}.`;
        if (parts.join("\n\n").length + next.length + 2 > MAX_CONTEXT_CHARS) break;
        parts.push(next);
      }
  }
  for (const entry of skills) {
    const next = `Skill "${entry.name}":\n${entry.instructions}`;
    if (parts.join("\n\n").length + next.length + 2 > MAX_CONTEXT_CHARS) break;
    parts.push(next);
  }
  return parts.join("\n\n");
}
