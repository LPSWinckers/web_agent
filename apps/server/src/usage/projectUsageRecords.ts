import {
  TurnTokenUsage,
  UsageProviderKind,
  type OrchestrationThreadActivity,
} from "@t3tools/contracts";
import * as Option from "effect/Option";
import * as Schema from "effect/Schema";

import type { UsageRecord } from "./usageTranscripts.ts";

const ProjectTurnUsage = Schema.Struct({
  provider: UsageProviderKind,
  model: Schema.String,
  tokenUsage: Schema.optional(TurnTokenUsage),
  reportedCostUsd: Schema.optional(Schema.Number),
});
const decode = Schema.decodeUnknownOption(ProjectTurnUsage);

export function projectUsageRecords(
  activities: ReadonlyArray<OrchestrationThreadActivity>,
): ReadonlyArray<UsageRecord> {
  return activities.flatMap((activity) => {
    if (activity.kind !== "turn.usage") return [];
    const payload = decode(activity.payload);
    if (Option.isNone(payload)) return [];
    const { provider, model, reportedCostUsd } = payload.value;
    if (!model.trim()) return [];
    const tokenUsage = payload.value.tokenUsage;
    if (
      tokenUsage?.inputTokens === undefined &&
      tokenUsage?.outputTokens === undefined &&
      reportedCostUsd === undefined
    )
      return [];
    const input = tokenUsage?.inputTokens ?? 0;
    const cached = Math.min(tokenUsage?.cachedInputTokens ?? 0, input);
    const cacheCreation = Math.min(tokenUsage?.cacheCreationTokens ?? 0, input - cached);
    const timestampMs = Date.parse(activity.createdAt);
    if (!Number.isFinite(timestampMs)) return [];
    return [
      {
        provider,
        model,
        timestampMs,
        sessionId: activity.turnId ?? "",
        totals: {
          uncachedInputTokens: input - cached - cacheCreation,
          cachedInputTokens: cached,
          cacheCreationTokens: cacheCreation,
          outputTokens: tokenUsage?.outputTokens ?? 0,
          reasoningTokens: tokenUsage?.reasoningTokens ?? 0,
        },
        reportedCostUsd: reportedCostUsd ?? null,
        fast: false,
        dedupeKey: activity.id,
      },
    ];
  });
}
