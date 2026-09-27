/* oxlint-disable react/no-array-index-key -- Chart labels can repeat; their ordered position identifies each point. */
import type { ConsultancyChart } from "@t3tools/shared/consultancyChart";
import { CONSULTANCY_PRESENTATION_STANDARD } from "@t3tools/shared/consultancyPresentationStandard";

const { colors, style } = CONSULTANCY_PRESENTATION_STANDARD;
const palette = [colors.blue, colors.navy, colors.cyan, colors.gold];

export function ConsultancyChartView({
  chart,
  compact = false,
}: {
  chart: ConsultancyChart;
  compact?: boolean;
}) {
  const format = (value: number) =>
    `${new Intl.NumberFormat("nl-NL", { maximumFractionDigits: 2 }).format(value)}${chart.unit === "%" ? "%" : chart.unit ? ` ${chart.unit}` : ""}`;
  const description = chart.series
    .map(
      (series) =>
        `${series.name}: ${chart.categories.map((category, index) => `${category} ${format(series.values[index]!)}`).join(", ")}`,
    )
    .join("; ");
  const allValues = chart.series.flatMap((series) => series.values);
  const minimum = Math.min(0, ...allValues);
  const maximum = Math.max(1, ...allValues);
  const scaleY = (value: number) => 128 - ((value - minimum) / (maximum - minimum || 1)) * 105;
  const scaleX = (index: number) => 30 + (index * 350) / Math.max(1, chart.categories.length - 1);

  let visual: React.ReactNode;
  if (chart.type === "bar") {
    const barMinimum = Math.min(0, ...allValues);
    const barMaximum = Math.max(0, ...allValues);
    const range = Math.max(1, barMaximum - barMinimum);
    const zero = (-barMinimum / range) * 100;
    visual = (
      <div className="space-y-2" role="img" aria-label={description}>
        {chart.categories.map((category, index) => (
          <div
            key={`${category}-${index}`}
            className="grid grid-cols-[minmax(0,8rem)_minmax(0,1fr)] items-center gap-2 text-xs"
          >
            <span className="min-w-0 truncate">{category}</span>
            <div className="space-y-1">
              {chart.series.map((series, seriesIndex) => (
                <div key={`${series.name}-${seriesIndex}`} className="flex items-center gap-2">
                  <div
                    className="relative h-2.5 flex-1 rounded-sm"
                    style={{ backgroundColor: `#${colors.paleBlue}` }}
                  >
                    {barMinimum < 0 ? (
                      <span
                        className="absolute inset-y-0 w-px bg-current opacity-40"
                        style={{ left: `${zero}%` }}
                      />
                    ) : null}
                    <div
                      className="absolute h-full rounded-sm"
                      style={{
                        left: `${series.values[index]! < 0 ? zero - (Math.abs(series.values[index]!) / range) * 100 : zero}%`,
                        width: `${(Math.abs(series.values[index]!) / range) * 100}%`,
                        backgroundColor: `#${palette[seriesIndex]}`,
                      }}
                    />
                  </div>
                  <span className="w-16 text-right tabular-nums">
                    {format(series.values[index]!)}
                  </span>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  } else if (chart.type === "pie" || chart.type === "doughnut") {
    const values = chart.series[0]!.values;
    const total = values.reduce((sum, value) => sum + value, 0);
    let progress = 0;
    const gradient = values.map((value, index) => {
      const start = progress;
      progress += total > 0 ? (value / total) * 100 : 0;
      return `#${palette[index % palette.length]} ${start}% ${progress}%`;
    });
    visual = (
      <div className="flex flex-wrap items-center gap-4" role="img" aria-label={description}>
        <div
          className="relative size-36 shrink-0 rounded-full"
          style={{
            background:
              total > 0 ? `conic-gradient(${gradient.join(", ")})` : `#${colors.paleBlue}`,
          }}
        >
          {chart.type === "doughnut" ? (
            <div
              className="absolute inset-8 rounded-full"
              style={{ backgroundColor: `#${style.background}` }}
            />
          ) : null}
        </div>
        <div className="space-y-1.5 text-xs">
          {chart.categories.map((category, index) => (
            <div key={`${category}-${index}`} className="flex items-center gap-2">
              <span
                className="size-2.5 shrink-0 rounded-sm"
                style={{ backgroundColor: `#${palette[index % palette.length]}` }}
              />
              <span>{category}</span>
              <span className="font-medium tabular-nums">{format(values[index]!)}</span>
            </div>
          ))}
        </div>
      </div>
    );
  } else {
    const scatter = chart.type === "scatter";
    const xValues = scatter ? chart.series[0]!.values : [];
    const yValues = scatter ? chart.series[1]!.values : [];
    const xMin = scatter ? Math.min(...xValues) : 0;
    const xMax = scatter ? Math.max(...xValues) : 1;
    const yMin = scatter ? Math.min(...yValues) : 0;
    const yMax = scatter ? Math.max(...yValues) : 1;
    const scatterX = (value: number) => 30 + ((value - xMin) / (xMax - xMin || 1)) * 350;
    const scatterY = (value: number) => 128 - ((value - yMin) / (yMax - yMin || 1)) * 105;
    visual = (
      <div role="img" aria-label={description}>
        <svg viewBox="0 0 410 150" className="h-36 w-full" aria-hidden="true">
          <line x1="30" x2="380" y1="128" y2="128" stroke={`#${style.muted}`} />
          <line x1="30" x2="30" y1="23" y2="128" stroke={`#${style.muted}`} />
          {scatter
            ? yValues.map((value, index) => (
                <circle
                  key={index}
                  cx={scatterX(xValues[index]!)}
                  cy={scatterY(value)}
                  r="5"
                  fill={`#${colors.blue}`}
                />
              ))
            : chart.series.map((series, index) => {
                const points = series.values
                  .map((value, pointIndex) => `${scaleX(pointIndex)},${scaleY(value)}`)
                  .join(" ");
                return (
                  <g key={`${series.name}-${index}`}>
                    {chart.type === "area" ? (
                      <polygon
                        points={`30,128 ${points} 380,128`}
                        fill={`#${palette[index]}`}
                        opacity="0.15"
                      />
                    ) : null}
                    <polyline
                      points={points}
                      fill="none"
                      stroke={`#${palette[index]}`}
                      strokeWidth="3"
                    />
                    {series.values.map((value, pointIndex) => (
                      <circle
                        key={pointIndex}
                        cx={scaleX(pointIndex)}
                        cy={scaleY(value)}
                        r="3.5"
                        fill={`#${palette[index]}`}
                      />
                    ))}
                  </g>
                );
              })}
        </svg>
        <div className="flex flex-wrap gap-x-3 gap-y-1 text-xs" aria-hidden="true">
          {chart.categories.map((category, index) => (
            <span key={`${category}-${index}`}>{category}</span>
          ))}
        </div>
      </div>
    );
  }

  return (
    <div
      className={`min-w-0 rounded-lg border ${compact ? "p-3" : "p-4"}`}
      style={{
        fontFamily: style.fontFace,
        borderColor: "#D8E1EA",
        backgroundColor: `#${style.background}`,
        color: `#${style.foreground}`,
      }}
      data-consultancy-chart={chart.type}
    >
      {chart.title ? <h4 className="mb-1 text-sm font-semibold">{chart.title}</h4> : null}
      {chart.series.length > 1 ? (
        <div className="mb-2 flex flex-wrap gap-3 text-xs">
          {chart.series.map((series, index) => (
            <span key={`${series.name}-${index}`}>
              <span style={{ color: `#${palette[index]}` }}>●</span> {series.name}
            </span>
          ))}
        </div>
      ) : null}
      {visual}
      {chart.source ? (
        <p className="mt-2 text-xs" style={{ color: `#${style.muted}` }}>
          Bron: {chart.source}
        </p>
      ) : null}
    </div>
  );
}
