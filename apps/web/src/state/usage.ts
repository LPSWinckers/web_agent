/**
 * Multi-environment usage state.
 *
 * Every connected environment answers the same typed query; the client merges
 * the results. Raw transcripts never leave the machine that produced them.
 *
 * @module state/usage
 */
import { useAtomValue } from "@effect/atom-react";
import {
  USAGE_CONTRACT_VERSION,
  type EnvironmentId,
  type UsageSummary,
  type UsageSummaryInput,
} from "@t3tools/contracts";
import type { EnvironmentProject } from "@t3tools/client-runtime/state/shell";
import { refreshUsage } from "@t3tools/client-runtime/state/usage";
import { executeAtomQuery } from "@t3tools/client-runtime/state/runtime";
import * as Option from "effect/Option";
import { AsyncResult, Atom } from "effect/unstable/reactivity";
import { useCallback, useMemo } from "react";

import { mergeUsage, type EnvironmentUsage, type MergedUsage } from "@t3tools/shared/usageMerge";
import { appAtomRegistry } from "../rpc/atomRegistry";
import { environmentProjects } from "./projects";
import { environmentPresentations } from "./presentation";
import { serverEnvironment } from "./server";

export interface EnvironmentUsageStatus {
  readonly environmentId: EnvironmentId;
  readonly label: string;
  readonly isPending: boolean;
  readonly error: string | null;
  readonly summary: UsageSummary | null;
}

/**
 * Reads every environment's summary for one window.
 *
 * Keyed by the serialised window so switching ranges does not thrash the atom
 * cache, and so each environment's query is shared with any other reader of the
 * same window.
 */
const usageByWindowAtom = Atom.family((windowKey: string) =>
  Atom.make((get): readonly EnvironmentUsageStatus[] => {
    const input = JSON.parse(windowKey) as UsageSummaryInput;
    const presentations = get(environmentPresentations.presentationsAtom);

    const statuses: EnvironmentUsageStatus[] = [];
    for (const [environmentId, presentation] of presentations) {
      const result = get(serverEnvironment.usageSummary({ environmentId, input }));
      statuses.push({
        environmentId,
        label: presentation.entry.target.label,
        isPending: result.waiting,
        error: result._tag === "Failure" ? "This environment could not report usage." : null,
        summary: Option.getOrNull(AsyncResult.value(result)),
      });
    }
    return statuses;
  }).pipe(Atom.withLabel(`web-usage:window:${windowKey}`)),
);

export interface EnvironmentProjectUsageStatus {
  readonly project: EnvironmentProject;
  readonly environmentLabel: string;
  readonly isPending: boolean;
  readonly error: string | null;
  readonly summary: UsageSummary | null;
}

const projectUsageByWindowAtom = Atom.family((selectionKey: string) =>
  Atom.make((get): readonly EnvironmentProjectUsageStatus[] => {
    const selection = JSON.parse(selectionKey) as {
      readonly inputKey: string;
      readonly environmentIds: readonly EnvironmentId[] | null;
      readonly enabled: boolean;
    };
    if (!selection.enabled) return [];

    const input = JSON.parse(selection.inputKey) as UsageSummaryInput;
    const presentations = get(environmentPresentations.presentationsAtom);
    return get(environmentProjects.projectsAtom)
      .filter(
        (project) =>
          selection.environmentIds === null ||
          selection.environmentIds.includes(project.environmentId),
      )
      .map((project) => {
        const result = get(
          serverEnvironment.usageSummary({
            environmentId: project.environmentId,
            input: { ...input, projectId: project.id },
          }),
        );
        return {
          project,
          environmentLabel:
            presentations.get(project.environmentId)?.entry.target.label ?? project.environmentId,
          isPending: result.waiting,
          error: result._tag === "Failure" ? "This project could not report usage." : null,
          summary: Option.getOrNull(AsyncResult.value(result)),
        };
      });
  }).pipe(Atom.withLabel(`web-usage:projects:${selectionKey}`)),
);

export interface UsageView {
  readonly merged: MergedUsage;
  readonly environments: readonly EnvironmentUsageStatus[];
  readonly selectedEnvironments: readonly EnvironmentUsageStatus[];
  /** True until at least one selected environment has answered. */
  readonly isPending: boolean;
  /**
   * True while environments that have not failed are still answering. Failed
   * environments are reported through their own error rows: totals will not
   * improve by waiting on them, so they must not read as "still reporting".
   */
  readonly isPartial: boolean;
  readonly refresh: (input?: UsageSummaryInput) => Promise<void>;
}

export function useUsage(
  input: UsageSummaryInput,
  selectedEnvironmentIds: ReadonlySet<EnvironmentId> | null = null,
): UsageView {
  const windowKey = useMemo(
    () =>
      JSON.stringify({
        sinceDay: input.sinceDay,
        untilDay: input.untilDay,
        timeZone: input.timeZone,
        resolution: input.resolution,
        sinceTime: input.sinceTime,
        untilTime: input.untilTime,
      }),
    [
      input.sinceDay,
      input.untilDay,
      input.timeZone,
      input.resolution,
      input.sinceTime,
      input.untilTime,
    ],
  );
  const atom = usageByWindowAtom(windowKey);
  const environments = useAtomValue(atom);
  const selectedEnvironments = useMemo(
    () =>
      selectedEnvironmentIds === null
        ? environments
        : environments.filter((environment) =>
            selectedEnvironmentIds.has(environment.environmentId),
          ),
    [environments, selectedEnvironmentIds],
  );

  const refresh = useCallback(
    (nextInput?: UsageSummaryInput) =>
      refreshUsage({
        registry: appAtomRegistry,
        server: serverEnvironment,
        presentations: environmentPresentations,
        environmentIds: selectedEnvironments.map(({ environmentId }) => environmentId),
        input: nextInput ?? (JSON.parse(windowKey) as UsageSummaryInput),
      }),
    [selectedEnvironments, windowKey],
  );

  const merged = useMemo(() => {
    const answered: EnvironmentUsage[] = selectedEnvironments.flatMap((environment) =>
      environment.summary === null
        ? []
        : [
            {
              environmentId: environment.environmentId,
              label: environment.label,
              summary: environment.summary,
            },
          ],
    );
    return mergeUsage(answered, USAGE_CONTRACT_VERSION);
  }, [selectedEnvironments]);

  const answeredCount = selectedEnvironments.filter(
    (environment) => environment.summary !== null,
  ).length;
  const stillReporting = selectedEnvironments.filter(
    (environment) => environment.summary === null && environment.error === null,
  ).length;

  return {
    merged,
    environments,
    selectedEnvironments,
    isPending: answeredCount === 0 && stillReporting > 0,
    isPartial: answeredCount > 0 && stillReporting > 0,
    refresh,
  };
}

export function useProjectUsage(
  input: UsageSummaryInput,
  selectedEnvironmentIds: ReadonlySet<EnvironmentId> | null = null,
  enabled = true,
) {
  const inputKey = useMemo(
    () =>
      JSON.stringify({
        sinceDay: input.sinceDay,
        untilDay: input.untilDay,
        timeZone: input.timeZone,
        resolution: input.resolution,
        sinceTime: input.sinceTime,
        untilTime: input.untilTime,
      }),
    [
      input.sinceDay,
      input.untilDay,
      input.timeZone,
      input.resolution,
      input.sinceTime,
      input.untilTime,
    ],
  );
  const environmentIds = useMemo(
    () => (selectedEnvironmentIds === null ? null : [...selectedEnvironmentIds].sort()),
    [selectedEnvironmentIds],
  );
  const selectionKey = JSON.stringify({ inputKey, environmentIds, enabled });
  const projects = useAtomValue(projectUsageByWindowAtom(selectionKey));

  const refresh = useCallback(
    async (nextInput?: UsageSummaryInput) => {
      const summaryInput = nextInput ?? (JSON.parse(inputKey) as UsageSummaryInput);
      await Promise.all(
        projects.map(async ({ project }) => {
          const query = serverEnvironment.usageSummary({
            environmentId: project.environmentId,
            input: { ...summaryInput, projectId: project.id },
          });
          appAtomRegistry.refresh(query);
          await executeAtomQuery(appAtomRegistry, query, { reportFailure: false });
        }),
      );
    },
    [inputKey, projects],
  );

  return {
    projects,
    isPending: projects.some((project) => project.isPending),
    failedCount: projects.filter((project) => project.error !== null).length,
    refresh,
  };
}
