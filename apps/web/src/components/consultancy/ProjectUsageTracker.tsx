import { UsageDay, type EnvironmentId, type ProjectId } from "@t3tools/contracts";
import { formatUsd } from "@t3tools/shared/usageFormat";
import { useEffect, useEffectEvent, useMemo } from "react";

import { Button } from "~/components/ui/button";
import { serverEnvironment } from "~/state/server";
import { useEnvironmentQuery } from "~/state/query";

export function ProjectUsageTracker({
  environmentId,
  projectId,
}: {
  environmentId: EnvironmentId;
  projectId: ProjectId;
}) {
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  const untilDay = UsageDay.make(new Date().toLocaleDateString("en-CA", { timeZone }));
  const input = useMemo(
    () => ({
      projectId,
      sinceDay: UsageDay.make("1970-01-01"),
      untilDay,
      timeZone,
    }),
    [projectId, timeZone, untilDay],
  );
  const usage = useEnvironmentQuery(serverEnvironment.usageSummary({ environmentId, input }));
  const refreshOnOpen = useEffectEvent(usage.refresh);
  useEffect(() => {
    refreshOnOpen();
  }, [environmentId, projectId]);

  const models = useMemo(() => {
    const rows = new Map<
      string,
      {
        provider: string;
        model: string;
        input: number;
        output: number;
        cost: number;
        unpriced: number;
      }
    >();
    for (const bucket of usage.data?.buckets ?? []) {
      const key = `${bucket.provider}\u0000${bucket.model}`;
      const row = rows.get(key) ?? {
        provider: bucket.provider,
        model: bucket.model,
        input: 0,
        output: 0,
        cost: 0,
        unpriced: 0,
      };
      row.input +=
        bucket.totals.uncachedInputTokens +
        bucket.totals.cachedInputTokens +
        bucket.totals.cacheCreationTokens;
      row.output += bucket.totals.outputTokens;
      row.cost += bucket.costUsd;
      row.unpriced += bucket.unpricedRecords;
      rows.set(key, row);
    }
    return [...rows.values()].sort((a, b) => b.input + b.output - a.input - a.output);
  }, [usage.data]);
  const totals = models.reduce(
    (sum, row) => ({
      tokens: sum.tokens + row.input + row.output,
      cost: sum.cost + row.cost,
      unpriced: sum.unpriced + row.unpriced,
    }),
    { tokens: 0, cost: 0, unpriced: 0 },
  );
  const count = (value: number) => value.toLocaleString();
  const cost = (value: number, unpriced: number) =>
    unpriced > 0 ? (value > 0 ? `${formatUsd(value)}+` : "Unpriced") : formatUsd(value);

  return (
    <section className="mt-8 rounded-xl border border-border bg-card p-5">
      <div className="flex items-start justify-between gap-3">
        <div>
          <h2 className="font-medium">Project AI usage</h2>
          <p className="mt-1 text-sm text-muted-foreground">
            Tokens and estimated API cost from this project&apos;s recorded AI turns.
          </p>
        </div>
        <Button variant="ghost" size="sm" onClick={usage.refresh} disabled={usage.isPending}>
          Refresh
        </Button>
      </div>
      {usage.error ? (
        <p role="alert" className="mt-4 text-sm text-destructive">
          {usage.error}
        </p>
      ) : null}
      {usage.isPending && !usage.data ? (
        <p className="mt-4 text-sm text-muted-foreground">Loading usage…</p>
      ) : models.length === 0 ? (
        <p className="mt-4 text-sm text-muted-foreground">No recorded AI usage yet.</p>
      ) : (
        <div className="mt-4 overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="border-b border-border text-left text-muted-foreground">
              <tr>
                <th className="py-2 font-medium">Model</th>
                <th className="px-3 py-2 text-right font-medium">Input</th>
                <th className="px-3 py-2 text-right font-medium">Output</th>
                <th className="px-3 py-2 text-right font-medium">Total tokens</th>
                <th className="py-2 text-right font-medium">Cost</th>
              </tr>
            </thead>
            <tbody>
              {models.map((row) => (
                <tr
                  key={`${row.provider}:${row.model}`}
                  className="border-b border-border/60 last:border-0"
                >
                  <td className="py-2">
                    <span className="font-medium">{row.model}</span>
                    <span className="ml-2 text-xs text-muted-foreground">{row.provider}</span>
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">{count(row.input)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{count(row.output)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {count(row.input + row.output)}
                  </td>
                  <td className="py-2 text-right tabular-nums">{cost(row.cost, row.unpriced)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {models.length > 0 ? (
        <p className="mt-4 text-sm font-medium">
          {count(totals.tokens)} tokens · {cost(totals.cost, totals.unpriced)} estimated API cost
        </p>
      ) : null}
      <p className="mt-2 text-xs text-muted-foreground">
        Counts reported for new turns by this server. Providers that omit usage and subscription
        charges are not included.
        {totals.unpriced > 0
          ? " Some models have no known price, so the cost shown is a minimum."
          : ""}
      </p>
    </section>
  );
}
