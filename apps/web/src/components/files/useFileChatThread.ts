import { scopeProjectRef, scopeThreadRef } from "@t3tools/client-runtime/environment";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";
import type { EnvironmentId, ProjectId, RuntimeMode, ScopedThreadRef } from "@t3tools/contracts";
import { useCallback, useEffect, useRef, useState } from "react";

import { useComposerDraftStore, type DraftId } from "~/composerDraftStore";
import { useNewThreadHandler } from "~/hooks/useHandleNewThread";
import { projectEnvironment } from "~/state/projects";
import { useAtomCommand } from "~/state/use-atom-command";
import { useAtomQueryRunner } from "~/state/use-atom-query-runner";

export function fileChatSidecarPath(path: string) {
  if (path.startsWith("browser:")) return `.werkbestanden/${encodeURIComponent(path)}.t3chat.json`;
  return `${path}.t3chat.json`;
}

export function parseFileChatSidecar(source: string, path: string) {
  try {
    const value: unknown = JSON.parse(source);
    if (
      typeof value === "object" &&
      value !== null &&
      "version" in value &&
      value.version === 1 &&
      "path" in value &&
      value.path === path &&
      "threadId" in value &&
      typeof value.threadId === "string" &&
      value.threadId.trim()
    ) {
      return value.threadId;
    }
  } catch {
    // A damaged sidecar is treated as missing when the file chat starts.
  }
  return null;
}

export function useFileChatThread(
  environmentId: EnvironmentId,
  projectId: ProjectId,
  cwd: string,
  path: string,
  runtimeMode: RuntimeMode,
) {
  const newThread = useNewThreadHandler();
  const readFile = useAtomQueryRunner(projectEnvironment.readFile, {
    reportFailure: false,
    refresh: true,
  });
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const [threadRef, setThreadRef] = useState<ScopedThreadRef | null>(null);
  const [draftId, setDraftId] = useState<DraftId | null>(null);
  const pending = useRef<Promise<{ threadRef: ScopedThreadRef; draftId: DraftId | null }> | null>(
    null,
  );
  const draftSession = useComposerDraftStore((store) =>
    draftId ? store.getDraftSession(draftId) : null,
  );

  useEffect(() => {
    const target = draftId && !draftSession?.promotedTo ? draftId : threadRef;
    if (target) useComposerDraftStore.getState().setRuntimeMode(target, runtimeMode);
  }, [draftId, draftSession?.promotedTo, runtimeMode, threadRef]);

  useEffect(() => {
    let cancelled = false;
    void readFile({
      environmentId,
      input: { cwd, relativePath: fileChatSidecarPath(path) },
    }).then((result) => {
      if (cancelled || result._tag !== "Success") return;
      const id = parseFileChatSidecar(result.value.contents, path);
      if (id) setThreadRef(scopeThreadRef(environmentId, id as ScopedThreadRef["threadId"]));
    });
    return () => {
      cancelled = true;
    };
  }, [cwd, environmentId, path, readFile]);

  const ensureThread = useCallback(() => {
    if (pending.current) return pending.current;
    const task = (async () => {
      const sidecar = await readFile({
        environmentId,
        input: { cwd, relativePath: fileChatSidecarPath(path) },
      });
      if (sidecar._tag === "Success") {
        const id = parseFileChatSidecar(sidecar.value.contents, path);
        if (id) {
          const ref = scopeThreadRef(environmentId, id as ScopedThreadRef["threadId"]);
          useComposerDraftStore.getState().setRuntimeMode(ref, runtimeMode);
          setThreadRef(ref);
          return { threadRef: ref, draftId: null };
        }
      }
      const created = await newThread(scopeProjectRef(environmentId, projectId), {
        navigate: false,
      });
      if (!created) throw new Error("Could not start the file chat.");
      useComposerDraftStore.getState().setRuntimeMode(created.draftId, runtimeMode);
      const contents = JSON.stringify({ version: 1, path, threadId: created.threadId });
      const saved = await writeFile({
        environmentId,
        input: { cwd, relativePath: fileChatSidecarPath(path), contents },
      });
      if (saved._tag === "Failure") throw squashAtomCommandFailure(saved);
      const ref = scopeThreadRef(environmentId, created.threadId);
      setThreadRef(ref);
      setDraftId(created.draftId);
      return { threadRef: ref, draftId: created.draftId };
    })();
    pending.current = task;
    const clearPending = () => {
      if (pending.current === task) pending.current = null;
    };
    void task.then(clearPending, clearPending);
    return task;
  }, [cwd, environmentId, newThread, path, projectId, readFile, runtimeMode, writeFile]);

  const target =
    draftId && draftSession && !draftSession.promotedTo
      ? ({ kind: "draft", draftId } as const)
      : threadRef
        ? ({ kind: "server", threadRef } as const)
        : null;
  return { ensureThread, target, threadRef };
}
