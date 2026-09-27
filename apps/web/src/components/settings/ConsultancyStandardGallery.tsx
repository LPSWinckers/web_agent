import { CONSULTANCY_CHART_TYPES, type ConsultancyChart } from "@t3tools/shared/consultancyChart";
import {
  CONSULTANCY_PRESENTATION_STANDARD,
  CONSULTANCY_SLIDE_COMPONENTS,
} from "@t3tools/shared/consultancyPresentationStandard";
import { ConsultancyChartView } from "../charts/ConsultancyChartView";

const { style, colors } = CONSULTANCY_PRESENTATION_STANDARD;

function exampleChart(type: ConsultancyChart["type"]): ConsultancyChart {
  return type === "scatter"
    ? {
        type,
        categories: ["A", "B", "C"],
        series: [
          { name: "Kosten", values: [2, 4, 6] },
          { name: "Impact", values: [3, 7, 8] },
        ],
      }
    : {
        type,
        categories: ["A", "B", "C"],
        series: [{ name: "Voorbeeld", values: [45, 70, 55] }],
        unit: "%",
      };
}

function SlidePreview({ type }: { type: (typeof CONSULTANCY_SLIDE_COMPONENTS)[number]["id"] }) {
  const dark = type === "cover" || type === "section";
  return (
    <div
      className={`relative aspect-video overflow-hidden rounded-md border p-4 shadow-sm ${dark ? "border-l-4" : "border-t-4"}`}
      style={{
        backgroundColor: `#${dark ? colors.navy : style.background}`,
        borderColor: `#${dark ? colors.gold : colors.navy}`,
        color: `#${dark ? style.background : style.foreground}`,
        fontFamily: style.fontFace,
      }}
    >
      <span className="text-3xs font-semibold tracking-wider opacity-75">
        BERENSCHOT 2026 / PRESENTATIE
      </span>
      <div className={`font-semibold ${dark ? "mt-7 text-xl" : "mt-4 text-lg"}`}>
        {type === "cover"
          ? "Titel van de presentatie"
          : type === "section"
            ? "Nieuwe sectie"
            : "Kernboodschap van de slide"}
      </div>
      <div className="mt-3 h-1 w-10" style={{ backgroundColor: `#${colors.gold}` }} />
      {type === "two-column" ? (
        <div className="mt-4 grid grid-cols-2 gap-2 text-3xs">
          <div className="rounded p-2" style={{ backgroundColor: `#${colors.paleBlue}` }}>
            INZICHT
            <br />
            Wat laten de gegevens zien?
          </div>
          <div className="rounded border p-2" style={{ borderColor: "#D8E1EA" }}>
            IMPLICATIE
            <br />
            Welke keuze volgt hieruit?
          </div>
        </div>
      ) : type === "chart" ? (
        <div className="mt-4 space-y-1.5" aria-hidden>
          {[78, 56, 39].map((width) => (
            <div
              key={width}
              className="h-2 rounded-sm"
              style={{ width: `${width}%`, backgroundColor: `#${colors.blue}` }}
            />
          ))}
        </div>
      ) : (
        <p className="mt-4 text-xs opacity-80">
          {dark
            ? "Korte ondertitel of overgang"
            : "Eén inzicht, een implicatie en een volgende stap."}
        </p>
      )}
      <span className="absolute bottom-2 left-4 text-3xs opacity-65">Bron: rapport of URL</span>
    </div>
  );
}

export function ConsultancyStandardGallery() {
  return (
    <>
      <section className="space-y-3 px-3 sm:px-4" aria-labelledby="presentation-standard-title">
        <div>
          <h2 id="presentation-standard-title" className="text-sm font-medium">
            Eén gedeelde PowerPoint-standaard
          </h2>
          <p className="mt-1 text-xs text-muted-foreground">
            {CONSULTANCY_PRESENTATION_STANDARD.name} is de standaard voor de maker,
            PowerPoint-export en AI. Dezelfde componenten zijn zichtbaar voor iedere gebruiker.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {[
            ["Marine", colors.navy],
            ["Blauw", colors.blue],
            ["Cyaan", colors.cyan],
            ["Goud", colors.gold],
            ["Lichtblauw", colors.paleBlue],
          ].map(([name, hex]) => (
            <div key={name} className="flex items-center gap-2 rounded-md border px-2 py-1 text-xs">
              <span className="size-4 rounded-sm border" style={{ backgroundColor: `#${hex}` }} />
              {name} <span className="text-muted-foreground">#{hex}</span>
            </div>
          ))}
        </div>
      </section>
      <section className="space-y-3 px-3 sm:px-4" aria-labelledby="slide-components-title">
        <h2 id="slide-components-title" className="text-sm font-medium">
          Slides
        </h2>
        <div className="grid gap-3 md:grid-cols-2">
          {CONSULTANCY_SLIDE_COMPONENTS.map((component) => (
            <article
              key={component.id}
              className="space-y-2 rounded-xl border border-border/60 bg-card/40 p-3"
            >
              <SlidePreview type={component.id} />
              <h3 className="text-sm font-medium">{component.name}</h3>
              <p className="text-xs text-muted-foreground">{component.description}</p>
            </article>
          ))}
        </div>
      </section>
      <section className="space-y-3 px-3 sm:px-4" aria-labelledby="chart-components-title">
        <h2 id="chart-components-title" className="text-sm font-medium">
          Modulaire grafieken
        </h2>
        <p className="text-xs text-muted-foreground">
          Eén staafgrafiek en vijf andere typen. De AI vult elk type met de cijfers en bron uit jouw
          gegevens.
        </p>
        <div className="grid items-start gap-3 md:grid-cols-2">
          {CONSULTANCY_CHART_TYPES.map((component) => (
            <article
              key={component.type}
              className="space-y-2 rounded-xl border border-border/60 bg-card/40 p-3"
            >
              <h3 className="text-sm font-medium">{component.name}</h3>
              <p className="text-xs text-muted-foreground">{component.description}</p>
              <ConsultancyChartView chart={exampleChart(component.type)} compact />
            </article>
          ))}
        </div>
      </section>
    </>
  );
}
