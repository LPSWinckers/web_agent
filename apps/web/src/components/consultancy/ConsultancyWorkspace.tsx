import { scopeProjectRef, scopeThreadRef } from "@t3tools/client-runtime/environment";
import type { ScopedThreadRef } from "@t3tools/contracts";
import { useAtomValue } from "@effect/atom-react";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";
import {
  FileSpreadsheetIcon,
  FileTextIcon,
  MessageSquareIcon,
  PlusIcon,
  XIcon,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";

import { SpreadsheetFileViewer } from "~/components/workbench/SpreadsheetFileViewer";
import { fileChatSidecarPath, parseFileChatSidecar } from "~/components/files/useFileChatThread";
import { ThreadRouteView } from "~/components/ThreadRouteView";
import { requestDraftAutoSend, type DraftId, useComposerDraftStore } from "~/composerDraftStore";
import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import { SidebarInset } from "~/components/ui/sidebar";
import { useNewThreadHandler } from "~/hooks/useHandleNewThread";
import { newProjectId } from "~/lib/utils";
import { useProjects, useThreadShell } from "~/state/entities";
import { usePrimaryEnvironmentId } from "~/state/environments";
import { projectEnvironment } from "~/state/projects";
import { useEnvironmentQuery } from "~/state/query";
import { useAtomCommand } from "~/state/use-atom-command";
import { useAtomQueryRunner } from "~/state/use-atom-query-runner";
import { filesystemEnvironment } from "~/state/filesystem";
import { environmentServerConfigsAtom } from "~/state/server";
import {
  ensureBrowseDirectoryPath,
  getBrowseParentPath,
} from "@t3tools/client-runtime/state/projects";
import {
  clearConsultancyProjectSelection,
  readActiveConsultancyProject,
  readPendingConsultancyCustomer,
  selectConsultancyProject,
  takePendingConsultancyCustomer,
} from "./ConsultancySidebar";
import {
  customerProjectTitle,
  deleteOriginalFile,
  parseManifest,
  readCustomers,
  readOriginalFile,
  splitCustomerProject,
  type ProjectDocument,
} from "./consultancyData";
import { ConsultancyFileViewer, ConsultancyProjectFiles } from "./ConsultancyProjectFiles";
import { ProjectUsageTracker } from "./ProjectUsageTracker";
import { projectFolderName, projectFolderPath } from "./consultancyProjectSetup";
import { consultancyProjectFolders, type ProjectSetupFile } from "./consultancyProjectFolders";
import {
  takeConsultancyThreadTabRequests,
  type ShowConsultancyOverviewDetail,
} from "./consultancyThreadTabs";
import {
  agentWorkspaceContext,
  defaultAgentWorkspaceProfiles,
  type AgentWorkspaceKind,
} from "../agentWorkspaces";

const MANIFEST_PATH = "consultancy/project.json";

type ProjectTab =
  | { kind: "file"; key: string; name: string; path: string }
  | { kind: "workbook"; key: string; name: string; file: File }
  | {
      kind: "chat";
      key: string;
      name: string;
      threadRef: ScopedThreadRef;
      draftId?: DraftId;
      agentKind: AgentWorkspaceKind;
      editorContext?: string;
      fileKey?: string;
    };

type ProjectTabs = { scope: string; open: ProjectTab[]; active: string | null };

function errorText(result: Parameters<typeof squashAtomCommandFailure>[0]): string {
  const error = squashAtomCommandFailure(result);
  return error instanceof Error ? error.message : "The request failed.";
}

function ProjectChatTab({ tab }: { tab: Extract<ProjectTab, { kind: "chat" }> }) {
  const thread = useThreadShell(tab.threadRef);
  return (
    <ThreadRouteView
      target={
        tab.draftId && !thread
          ? { kind: "draft", draftId: tab.draftId }
          : { kind: "server", threadRef: tab.threadRef }
      }
      embedded
      {...(tab.editorContext ? { editorContext: tab.editorContext } : {})}
    />
  );
}

export function ConsultancyWorkspace() {
  const projects = useProjects();
  const environmentServerConfigs = useAtomValue(environmentServerConfigsAtom);
  const environmentId = usePrimaryEnvironmentId();
  const createProject = useAtomCommand(projectEnvironment.create, { reportFailure: false });
  const deleteProject = useAtomCommand(projectEnvironment.delete, { reportFailure: false });
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const readFile = useAtomQueryRunner(projectEnvironment.readFile, {
    reportFailure: false,
    refresh: true,
  });
  const browseFolders = useAtomQueryRunner(filesystemEnvironment.browse, { reportFailure: false });
  const newThread = useNewThreadHandler();
  const [selectedId, setSelectedId] = useState(readActiveConsultancyProject);
  const [customer, setCustomer] = useState(readPendingConsultancyCustomer);
  const [projectName, setProjectName] = useState("");
  const [folder, setFolder] = useState("");
  const [browsePath, setBrowsePath] = useState<string | null>(null);
  const [showCreate, setShowCreate] = useState(false);
  const [busy, setBusy] = useState(false);
  const [pendingSetup, setPendingSetup] = useState<{
    projectId: ReturnType<typeof newProjectId>;
    root: string;
    files: ProjectSetupFile[];
    next: number;
  } | null>(null);
  const [error, setError] = useState("");
  const [tabState, setTabState] = useState<ProjectTabs>({ scope: "", open: [], active: null });
  const [dirtyFileTabs, setDirtyFileTabs] = useState<ReadonlySet<string>>(() => new Set());
  const [documentsOverride, setDocumentsOverride] = useState<ReadonlyArray<ProjectDocument> | null>(
    null,
  );
  const [customersVersion, setCustomersVersion] = useState(0);
  const project = projects.find((entry) => entry.id === selectedId) ?? null;
  const agentProfiles = defaultAgentWorkspaceProfiles(
    project
      ? environmentServerConfigs.get(project.environmentId)?.settings.agentWorkspaces
      : undefined,
  );
  const projectScope = project ? `${project.environmentId}:${project.id}` : "";
  const tabs = tabState.scope === projectScope ? tabState.open : [];
  const activeTab = tabState.scope === projectScope ? tabState.active : null;
  const projectIdentity = project ? splitCustomerProject(project) : null;
  const customers = useMemo(
    () =>
      [
        ...new Set([
          ...readCustomers(),
          ...projects
            .map((entry) => splitCustomerProject(entry)?.customer)
            .filter((name): name is string => !!name),
        ]),
      ].sort(),
    [projects, customersVersion],
  );
  const manifestQuery = useEnvironmentQuery(
    project
      ? projectEnvironment.readFile({
          environmentId: project.environmentId,
          input: { cwd: project.workspaceRoot, relativePath: MANIFEST_PATH },
        })
      : null,
  );
  const manifest = useMemo(
    () => parseManifest(manifestQuery.data?.contents),
    [manifestQuery.data?.contents],
  );
  const documents = documentsOverride ?? manifest.documents;
  const proposedFolder = useMemo(() => {
    if (!folder.trim() || !projectName.trim()) return null;
    try {
      return projectFolderPath(folder, projectName, new Date());
    } catch {
      return null;
    }
  }, [folder, projectName]);
  const folderQuery = useEnvironmentQuery(
    browsePath && environmentId
      ? filesystemEnvironment.browse({
          environmentId,
          input: { partialPath: browsePath },
        })
      : null,
  );

  useEffect(() => {
    const update = () => {
      setSelectedId(readActiveConsultancyProject());
      setDocumentsOverride(null);
      setTabState({ scope: "", open: [], active: null });
    };
    window.addEventListener("consultancy-project-change", update);
    return () => window.removeEventListener("consultancy-project-change", update);
  }, []);

  useEffect(() => {
    const showOverview = (event: Event) => {
      const { detail } = event as CustomEvent<ShowConsultancyOverviewDetail>;
      const scope = `${detail.environmentId}:${detail.projectId}`;
      if (readActiveConsultancyProject() !== detail.projectId) {
        selectConsultancyProject(detail.projectId);
      }
      setSelectedId(detail.projectId);
      setTabState((previous) => ({
        scope,
        open: previous.scope === scope ? previous.open : [],
        active: null,
      }));
    };
    window.addEventListener("consultancy-show-overview", showOverview);
    return () => window.removeEventListener("consultancy-show-overview", showOverview);
  }, []);

  useEffect(() => {
    const openQueuedThreads = () => {
      for (const detail of takeConsultancyThreadTabRequests()) {
        const key = `chat:${detail.threadRef.environmentId}:${detail.threadRef.threadId}`;
        const detailProjectScope = `${detail.threadRef.environmentId}:${detail.projectId}`;
        setSelectedId(detail.projectId);
        setTabState((previous) => {
          const open = previous.scope === detailProjectScope ? previous.open : [];
          const existing = open.some((tab) => tab.key === key);
          return {
            scope: detailProjectScope,
            open: existing
              ? open
              : [
                  ...open,
                  {
                    kind: "chat",
                    key,
                    name: detail.title,
                    threadRef: detail.threadRef,
                    agentKind: "general",
                  },
                ],
            active: key,
          };
        });
      }
    };
    window.addEventListener("consultancy-open-thread", openQueuedThreads);
    openQueuedThreads();
    return () => window.removeEventListener("consultancy-open-thread", openQueuedThreads);
  }, []);

  useEffect(() => {
    const pending = takePendingConsultancyCustomer();
    if (pending) {
      setCustomer(pending);
      setShowCreate(true);
    }
    const update = () => {
      setCustomer(takePendingConsultancyCustomer());
      setShowCreate(true);
    };
    window.addEventListener("consultancy-customer-change", update);
    return () => window.removeEventListener("consultancy-customer-change", update);
  }, []);

  const write = async (path: string, contents: string) => {
    if (!project) throw new Error("Select a project first.");
    const result = await writeFile({
      environmentId: project.environmentId,
      input: { cwd: project.workspaceRoot, relativePath: path, contents },
    });
    if (result._tag === "Failure") throw new Error(errorText(result));
  };

  const create = async () => {
    if (
      !environmentId ||
      (!pendingSetup && (!customer.trim() || !projectName.trim() || !folder.trim()))
    )
      return;
    setBusy(true);
    setError("");
    try {
      let setup = pendingSetup;
      if (!setup) {
        const now = new Date();
        const root = projectFolderPath(folder, projectName, now);
        const name = projectFolderName(projectName, now);
        const parent = await browseFolders({
          environmentId,
          input: { partialPath: ensureBrowseDirectoryPath(folder.trim()) },
        });
        if (parent._tag === "Failure") throw new Error("Choose an existing parent folder.");
        if (parent.value.entries.some((entry) => entry.name.toLowerCase() === name.toLowerCase())) {
          throw new Error(`A folder named ${name} already exists there.`);
        }
        const files = consultancyProjectFolders();
        const projectId = newProjectId();
        const result = await createProject({
          environmentId,
          input: {
            projectId,
            title: customerProjectTitle(customer, projectName),
            workspaceRoot: root,
            createWorkspaceRootIfMissing: true,
            defaultModelSelection: null,
          },
        });
        if (result._tag === "Failure") throw new Error(errorText(result));
        setup = { projectId, root, files, next: 0 };
        setPendingSetup(setup);
        selectConsultancyProject(projectId);
        setSelectedId(projectId);
      }
      for (let index = setup.next; index < setup.files.length; index++) {
        const file = setup.files[index]!;
        const saved = await writeFile({
          environmentId,
          input: {
            cwd: setup.root,
            relativePath: file.path,
            contents: file.contents,
          },
        });
        if (saved._tag === "Failure") {
          setPendingSetup({ ...setup, next: index });
          throw new Error(
            `Project folder was created, but ${file.path} could not be saved. ${errorText(saved)}`,
          );
        }
      }
      setPendingSetup(null);
      selectConsultancyProject(setup.projectId);
      setSelectedId(setup.projectId);
      setShowCreate(false);
      setProjectName("");
      setFolder("");
      setBrowsePath(null);
      setCustomersVersion((version) => version + 1);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not create the project.");
    } finally {
      setBusy(false);
    }
  };

  const removeDocument = async (document: ProjectDocument) => {
    if (!project || !window.confirm(`Remove ${document.name} from this project?`)) return;
    setBusy(true);
    setError("");
    try {
      const next = documents.filter((entry) => entry.id !== document.id);
      await write(MANIFEST_PATH, JSON.stringify({ version: 1, documents: next }, null, 2));
      await write(
        "CONSULTANCY_CONTEXT.md",
        `# Project context\n\nCustomer: ${projectIdentity?.customer ?? ""}\nProject: ${projectIdentity?.name ?? ""}\n\nRead the relevant files before answering.\n\n${next.map((entry) => `- ${entry.name}: ${entry.contextPath}`).join("\n")}\n`,
      );
      await write(document.contextPath, "This document has been removed from the project.\n");
      await deleteOriginalFile(`${project.environmentId}:${project.id}:${document.id}`);
      setDocumentsOverride(next);
      manifestQuery.refresh();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not remove the file.");
    } finally {
      setBusy(false);
    }
  };

  const openSpreadsheet = async (document: ProjectDocument) => {
    if (!project) return;
    const file = await readOriginalFile(`${project.environmentId}:${project.id}:${document.id}`);
    if (file) {
      const key = `workbook:${document.id}`;
      setTabState((previous) => ({
        scope: projectScope,
        open:
          previous.scope === projectScope && previous.open.some((tab) => tab.key === key)
            ? previous.open
            : [
                ...(previous.scope === projectScope ? previous.open : []),
                { kind: "workbook", key, name: file.name, file },
              ],
        active: key,
      }));
    } else setError("This workbook is stored in the browser where it was added.");
  };

  const openFile = useCallback(
    (path: string) => {
      if (!projectScope) return;
      const key = `file:${path}`;
      setTabState((previous) => ({
        scope: projectScope,
        open:
          previous.scope === projectScope && previous.open.some((tab) => tab.key === key)
            ? previous.open
            : [
                ...(previous.scope === projectScope ? previous.open : []),
                { kind: "file", key, name: path.split("/").at(-1) ?? path, path },
              ],
        active: key,
      }));
    },
    [projectScope],
  );

  const closeTab = (key: string) => {
    const dirtyKey = `${projectScope}:${key}`;
    if (
      dirtyFileTabs.has(dirtyKey) &&
      !window.confirm("Close this document and discard unsaved changes?")
    )
      return;
    setDirtyFileTabs((previous) => {
      if (!previous.has(dirtyKey)) return previous;
      const next = new Set(previous);
      next.delete(dirtyKey);
      return next;
    });
    setTabState((previous) => {
      if (previous.scope !== projectScope) return previous;
      const index = previous.open.findIndex((tab) => tab.key === key);
      if (index < 0) return previous;
      const open = previous.open.filter((tab) => tab.key !== key);
      return {
        ...previous,
        open,
        active:
          previous.active === key
            ? (open[Math.min(index, open.length - 1)]?.key ?? null)
            : previous.active,
      };
    });
  };

  const openAgentChat = async (input: {
    kind: AgentWorkspaceKind;
    title: string;
    agentContext: string;
    prompt?: string;
    fileKey?: string;
    activate?: boolean;
  }) => {
    if (!project) return false;
    try {
      const { kind, title, agentContext, prompt, fileKey } = input;
      const persist =
        kind !== "application" && fileKey !== undefined && !fileKey.startsWith("browser:");
      const existingTab = tabs.find(
        (tab) => tab.kind === "chat" && tab.agentKind === kind && tab.fileKey === fileKey,
      );
      const sidecar =
        fileKey !== undefined && persist
          ? await readFile({
              environmentId: project.environmentId,
              input: { cwd: project.workspaceRoot, relativePath: fileChatSidecarPath(fileKey) },
            })
          : null;
      const sidecarId =
        sidecar?._tag === "Success" && fileKey !== undefined
          ? parseFileChatSidecar(sidecar.value.contents, fileKey)
          : null;
      const existingId = existingTab?.kind === "chat" ? existingTab.threadRef.threadId : sidecarId;
      const result = existingId
        ? null
        : await newThread(scopeProjectRef(project.environmentId, project.id), { navigate: false });
      if (!existingId && !result) throw new Error(`Could not start the ${title.toLowerCase()}.`);
      if (persist && fileKey !== undefined && result) {
        const saved = await writeFile({
          environmentId: project.environmentId,
          input: {
            cwd: project.workspaceRoot,
            relativePath: fileChatSidecarPath(fileKey),
            contents: JSON.stringify({ version: 1, path: fileKey, threadId: result.threadId }),
          },
        });
        if (saved._tag === "Failure") throw errorText(saved);
      }
      const threadRef = scopeThreadRef(
        project.environmentId,
        (existingId ?? result!.threadId) as ScopedThreadRef["threadId"],
      );
      const runtimeMode =
        kind === "application" ? "approval-required" : agentProfiles[kind].runtimeMode;
      useComposerDraftStore.getState().setRuntimeMode(result?.draftId ?? threadRef, runtimeMode);
      if (prompt) {
        requestDraftAutoSend(result?.draftId ?? threadRef, prompt);
        useComposerDraftStore.getState().setPrompt(result?.draftId ?? threadRef, prompt);
      }
      const key = `chat:${project.environmentId}:${threadRef.threadId}`;
      setTabState((previous) => {
        const open = previous.scope === projectScope ? previous.open : [];
        return {
          scope: projectScope,
          open: open.some((tab) => tab.key === key)
            ? open.map((tab) =>
                tab.key === key && tab.kind === "chat"
                  ? { ...tab, agentKind: kind, editorContext: agentContext, name: title }
                  : tab,
              )
            : [
                ...open,
                {
                  kind: "chat",
                  key,
                  name: title,
                  threadRef,
                  agentKind: kind,
                  ...(result ? { draftId: result.draftId } : {}),
                  editorContext: agentContext,
                  ...(fileKey !== undefined ? { fileKey } : {}),
                },
              ],
          active: input.activate ? key : previous.active,
        };
      });
      return true;
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : `Could not start the ${input.title}.`);
      return false;
    }
  };

  const askProjectAi = (prompt: string) => {
    void openAgentChat({
      kind: "general",
      title: "General agent",
      agentContext: agentWorkspaceContext("general", agentProfiles.general),
      prompt,
      activate: true,
    });
    return Boolean(project);
  };

  const startChat = async () => {
    if (!project) return;
    await openAgentChat({
      kind: "general",
      title: "General agent",
      agentContext: agentWorkspaceContext("general", agentProfiles.general),
      activate: true,
    });
  };

  const askProjectSpreadsheet = (question: string, agentContext: string, path?: string) =>
    openAgentChat({
      kind: "excel",
      title: "Excel agent",
      agentContext,
      prompt: question,
      fileKey: path ?? `browser:${project?.id ?? "workbook"}`,
    });

  const startApplicationAgent = (path: string) =>
    openAgentChat({
      kind: "application",
      title: "Application agent",
      agentContext: agentWorkspaceContext(
        "application",
        agentProfiles.application,
        `Open file path: ${JSON.stringify(path)}. Read the file from the project workspace before answering.`,
      ),
      fileKey: path,
    });

  const agentChatTabFor = (fileKey: string, kind: AgentWorkspaceKind) => {
    const chat = tabs.find(
      (tab) => tab.kind === "chat" && tab.fileKey === fileKey && tab.agentKind === kind,
    );
    return chat?.kind === "chat" ? chat : null;
  };

  const agentChatFor = (fileKey: string, kind: AgentWorkspaceKind) => {
    const chat = agentChatTabFor(fileKey, kind);
    return chat ? <ProjectChatTab key={chat.key} tab={chat} /> : null;
  };

  const downloadOriginal = async (document: ProjectDocument) => {
    if (!project) return;
    const file = await readOriginalFile(`${project.environmentId}:${project.id}:${document.id}`);
    if (!file) {
      setError("The original file is saved in the browser where it was added.");
      return;
    }
    const url = URL.createObjectURL(file);
    const link = window.document.createElement("a");
    link.href = url;
    link.download = file.name;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
  };

  const removeProject = async () => {
    if (
      !project ||
      !window.confirm(
        `Remove ${projectIdentity?.name ?? "this project"} and its chats? Files in the project folder will remain on disk. This cannot be undone.`,
      )
    )
      return;
    setBusy(true);
    setError("");
    try {
      const result = await deleteProject({
        environmentId: project.environmentId,
        input: { projectId: project.id, force: true },
      });
      if (result._tag === "Failure") throw new Error(errorText(result));
      await Promise.allSettled(
        documents.map((document) =>
          deleteOriginalFile(`${project.environmentId}:${project.id}:${document.id}`),
        ),
      );
      clearConsultancyProjectSelection();
      setSelectedId(null);
      setDocumentsOverride(null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not remove the project.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <SidebarInset className="h-dvh min-h-0 overflow-hidden">
      {project && tabs.length > 0 ? (
        <div
          className="flex shrink-0 items-stretch overflow-x-auto border-b border-border bg-muted/30"
          role="tablist"
          aria-label={`${projectIdentity?.name ?? "Project"} screens`}
        >
          {tabs.map((tab) => (
            <div
              key={tab.key}
              className={`flex max-w-56 shrink-0 items-center border-r border-border ${activeTab === tab.key ? "border-b-2 border-b-primary bg-background" : "text-muted-foreground hover:bg-muted"}`}
            >
              <button
                type="button"
                role="tab"
                aria-selected={activeTab === tab.key}
                aria-label={`Open ${tab.kind === "file" ? tab.path : tab.name}`}
                className="min-w-0 truncate py-3 pl-4 pr-2 text-left text-sm"
                onClick={() => setTabState((previous) => ({ ...previous, active: tab.key }))}
              >
                {tab.name}
              </button>
              <button
                type="button"
                aria-label={`Close ${tab.name}`}
                className="mr-2 rounded p-1 hover:bg-muted"
                onClick={() => closeTab(tab.key)}
              >
                <XIcon className="size-3.5" />
              </button>
            </div>
          ))}
        </div>
      ) : null}
      {project &&
        tabs.map((tab) => (
          <div
            key={tab.key}
            role="tabpanel"
            aria-hidden={activeTab !== tab.key}
            className={activeTab === tab.key ? "flex min-h-0 flex-1 overflow-hidden" : "hidden"}
          >
            <div className="flex min-h-0 min-w-0 flex-1 flex-col">
              {tab.kind === "chat" ? (
                activeTab === tab.key ? (
                  <ProjectChatTab tab={tab} />
                ) : null
              ) : tab.kind === "file" ? (
                <ConsultancyFileViewer
                  environmentId={project.environmentId}
                  cwd={project.workspaceRoot}
                  path={tab.path}
                  active={activeTab === tab.key}
                  onAskAi={askProjectAi}
                  onAskSpreadsheet={askProjectSpreadsheet}
                  spreadsheetChat={activeTab === tab.key ? agentChatFor(tab.path, "excel") : null}
                  spreadsheetThreadRef={agentChatTabFor(tab.path, "excel")?.threadRef}
                  applicationChat={
                    activeTab === tab.key ? agentChatFor(tab.path, "application") : null
                  }
                  onStartApplicationAgent={() => void startApplicationAgent(tab.path)}
                  excelAgentProfile={agentProfiles.excel}
                  powerpointAgentProfile={agentProfiles.powerpoint}
                  projectId={project.id}
                  onDirtyChange={(dirty) =>
                    setDirtyFileTabs((previous) => {
                      const dirtyKey = `${projectScope}:${tab.key}`;
                      if (previous.has(dirtyKey) === dirty) return previous;
                      const next = new Set(previous);
                      if (dirty) next.add(dirtyKey);
                      else next.delete(dirtyKey);
                      return next;
                    })
                  }
                />
              ) : (
                <SpreadsheetFileViewer
                  file={tab.file}
                  inline
                  agentChat={
                    activeTab === tab.key ? agentChatFor(`browser:${tab.key}`, "excel") : null
                  }
                  workspaceProfile={agentProfiles.excel}
                  onAskAi={(question, agentContext) =>
                    askProjectSpreadsheet(question, agentContext, `browser:${tab.key}`)
                  }
                />
              )}
            </div>
          </div>
        ))}
      <div
        role={project && activeTab !== null ? "tabpanel" : undefined}
        aria-hidden={activeTab !== null}
        className={activeTab === null ? "min-h-0 flex-1 overflow-y-auto" : "hidden"}
      >
        <div className="mx-auto w-full max-w-5xl px-6 py-10">
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div>
              <p className="text-xs font-medium uppercase tracking-wide text-primary">
                Consultancy workspace
              </p>
              <h1 className="mt-2 text-3xl font-semibold">
                {projectIdentity?.name ?? "Customers and projects"}
              </h1>
              <p className="mt-2 text-sm text-muted-foreground">
                {projectIdentity
                  ? `${projectIdentity.customer} · Project files and AI conversations`
                  : "Create a customer project to keep its data, documents, and chats together."}
              </p>
            </div>
            <Button
              variant="outline"
              onClick={() => {
                setShowCreate(true);
                setCustomer(projectIdentity?.customer ?? customers[0] ?? "");
              }}
            >
              <PlusIcon /> New project
            </Button>
          </div>
          {error ? (
            <p
              role="alert"
              className="mt-6 rounded-md border border-destructive/50 bg-destructive/10 p-3 text-sm text-destructive"
            >
              {error}
            </p>
          ) : null}
          {showCreate || !project ? (
            <section className="mt-8 rounded-xl border border-border bg-card p-6">
              <h2 className="text-lg font-medium">Create project</h2>
              <p className="mt-1 text-sm text-muted-foreground">
                Choose a parent folder on the connected environment. A dated project folder will be
                created inside it.
              </p>
              <div className="mt-5 grid gap-3 sm:grid-cols-2">
                <Input
                  list="consultancy-customers"
                  placeholder="Customer name"
                  value={customer}
                  disabled={pendingSetup !== null}
                  onChange={(event) => setCustomer(event.target.value)}
                />
                <datalist id="consultancy-customers">
                  {customers.map((name) => (
                    <option key={name} value={name} />
                  ))}
                </datalist>
                <Input
                  placeholder="Project name"
                  value={projectName}
                  disabled={pendingSetup !== null}
                  onChange={(event) => setProjectName(event.target.value)}
                />
                <div className="flex gap-2 sm:col-span-2">
                  <Input
                    aria-label="Parent folder"
                    placeholder="Parent folder, e.g. G:\\Consultancy\\Acme"
                    value={folder}
                    disabled={pendingSetup !== null}
                    onChange={(event) => setFolder(event.target.value)}
                  />
                  <Button
                    type="button"
                    variant="outline"
                    disabled={pendingSetup !== null}
                    onClick={() => setBrowsePath(ensureBrowseDirectoryPath(folder || "~/"))}
                  >
                    Browse
                  </Button>
                </div>
              </div>
              {browsePath ? (
                <div className="mt-3 rounded-md border border-border p-3">
                  <div className="flex items-center gap-2">
                    <Input
                      aria-label="Browse path"
                      value={browsePath}
                      onChange={(event) => setBrowsePath(event.target.value)}
                    />
                    <Button
                      type="button"
                      variant="outline"
                      disabled={!getBrowseParentPath(browsePath)}
                      onClick={() => setBrowsePath(getBrowseParentPath(browsePath))}
                    >
                      Up
                    </Button>
                    <Button
                      type="button"
                      onClick={() => {
                        if (folderQuery.data) {
                          setFolder(folderQuery.data.parentPath);
                          setBrowsePath(null);
                        }
                      }}
                      disabled={!folderQuery.data}
                    >
                      Use folder
                    </Button>
                  </div>
                  <div className="mt-2 max-h-48 overflow-y-auto">
                    {folderQuery.data?.entries.map((entry) => (
                      <button
                        key={entry.fullPath}
                        type="button"
                        className="block w-full rounded px-2 py-1.5 text-left text-sm hover:bg-muted"
                        onClick={() => setBrowsePath(ensureBrowseDirectoryPath(entry.fullPath))}
                      >
                        {entry.name}
                      </button>
                    ))}
                    {folderQuery.error ? (
                      <p className="text-sm text-destructive">{folderQuery.error}</p>
                    ) : null}
                  </div>
                </div>
              ) : null}
              {proposedFolder ? (
                <p className="mt-3 text-xs text-muted-foreground">New folder: {proposedFolder}</p>
              ) : null}
              <div className="mt-4 flex gap-2">
                <Button
                  disabled={
                    busy ||
                    !environmentId ||
                    (!pendingSetup && (!customer.trim() || !projectName.trim() || !folder.trim()))
                  }
                  onClick={() => void create()}
                >
                  {pendingSetup ? "Finish setup" : "Create project"}
                </Button>
                {project ? (
                  <Button
                    variant="ghost"
                    disabled={pendingSetup !== null}
                    onClick={() => setShowCreate(false)}
                  >
                    Cancel
                  </Button>
                ) : null}
              </div>
            </section>
          ) : null}
          {project && !showCreate ? (
            <>
              <button
                type="button"
                onClick={() => void startChat()}
                className="w-full rounded-xl border border-border bg-card p-5 text-left hover:border-primary/60"
              >
                <MessageSquareIcon className="size-5 text-primary" />
                <h2 className="mt-3 font-medium">Ask AI about this project</h2>
                <p className="mt-1 text-sm text-muted-foreground">
                  Start a conversation about the files in this project.
                </p>
              </button>
              <ConsultancyProjectFiles
                key={`${project.environmentId}:${project.id}`}
                environmentId={project.environmentId}
                cwd={project.workspaceRoot}
                onOpenFile={openFile}
                onAskAi={askProjectAi}
                powerpointAgentProfile={agentProfiles.powerpoint}
                projectId={project.id}
              />
              {documents.length ? (
                <section className="mt-8">
                  <h2 className="text-lg font-medium">Previously added files</h2>
                  <div className="mt-4 overflow-hidden rounded-xl border border-border bg-card">
                    {documents.map((document) => (
                      <div
                        key={document.id}
                        className="flex items-center gap-3 border-b border-border px-4 py-3 last:border-b-0"
                      >
                        {document.kind === "spreadsheet" ? (
                          <FileSpreadsheetIcon className="size-4 text-primary" />
                        ) : (
                          <FileTextIcon className="size-4 text-primary" />
                        )}
                        <div className="min-w-0 flex-1">
                          <p className="truncate text-sm font-medium">{document.name}</p>
                          <p className="text-xs text-muted-foreground">
                            {Math.round(document.size / 1024)} KB · {document.kind}
                          </p>
                        </div>
                        {document.kind === "spreadsheet" ? (
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => void openSpreadsheet(document)}
                          >
                            Open
                          </Button>
                        ) : null}
                        <Button
                          size="sm"
                          variant="ghost"
                          onClick={() => void downloadOriginal(document)}
                        >
                          Download
                        </Button>
                        <Button
                          size="sm"
                          variant="ghost"
                          disabled={busy}
                          onClick={() => void removeDocument(document)}
                        >
                          Remove
                        </Button>
                      </div>
                    ))}
                  </div>
                </section>
              ) : null}
              <ProjectUsageTracker environmentId={project.environmentId} projectId={project.id} />
              <div className="mt-8 border-t border-border pt-5">
                <Button variant="ghost" disabled={busy} onClick={() => void removeProject()}>
                  Remove project
                </Button>
              </div>
            </>
          ) : null}
        </div>
      </div>
    </SidebarInset>
  );
}
