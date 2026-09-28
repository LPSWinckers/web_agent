import type {
  AgentWorkspaceProfileSettings,
  EnvironmentId,
  ProjectId,
  ScopedThreadRef,
} from "@t3tools/contracts";
import { ChevronLeftIcon, FolderIcon, FileSpreadsheetIcon, FileTextIcon } from "lucide-react";
import { lazy, Suspense, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";

import { useAssetUrlRefresh, useAssetUrlState } from "~/assets/assetUrls";
import ChatMarkdown from "~/components/ChatMarkdown";
import { BrowserDocumentFrame } from "~/components/files/BrowserDocumentFrame";
import { PresentationDeckPreview, PresentationMaker } from "~/components/files/PresentationMaker";
import { SpreadsheetFileViewer } from "~/components/workbench/SpreadsheetFileViewer";
import { workbookBase64 } from "~/components/workbench/workbookBase64";
import {
  exportWordDocument,
  newWordDocument,
  wordDocumentPath,
} from "~/components/files/wordDocument";
import { useDirectoryEntries } from "~/components/files/useDirectoryEntries";
import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import { useEnvironmentSettings } from "~/hooks/useSettings";
import {
  latestWorkspaceMutationId,
  useWorkspaceMutationRefresh,
} from "~/hooks/useWorkspaceMutationRefresh";
import { useThread } from "~/state/entities";
import { projectEnvironment } from "~/state/projects";
import { useEnvironmentQuery } from "~/state/query";
import { useAtomCommand } from "~/state/use-atom-command";
import { useAtomQueryRunner } from "~/state/use-atom-query-runner";

import { isConsultancyInternalPath } from "./consultancyData";

const SPREADSHEET = /\.(xlsx|csv|tsv)$/i;
const MARKDOWN = /\.md$/i;
const TEXT = /\.(md|txt|json|xml|html|csv|tsv)$/i;
const IMAGE = /\.(png|jpe?g|gif|webp|svg)$/i;
const WordDocumentEditor = lazy(() =>
  import("~/components/files/WordDocumentEditor").then((module) => ({
    default: module.WordDocumentEditor,
  })),
);

function ApplicationFileLayout({
  children,
  chat,
  onStart,
}: {
  children: ReactNode;
  chat?: ReactNode;
  onStart: () => void;
}) {
  return (
    <div className="flex min-h-0 flex-1 flex-col lg:flex-row">
      <main className="flex min-h-0 min-w-0 flex-1 flex-col overflow-auto">{children}</main>
      <aside className="flex min-h-52 flex-col border-t border-border lg:w-[min(42%,34rem)] lg:border-l lg:border-t-0">
        <div className="border-b border-border px-4 py-3">
          <h2 className="text-sm font-semibold">Application agent</h2>
          <p className="mt-1 text-xs text-muted-foreground">Read-only file questions</p>
        </div>
        {chat ?? (
          <div className="flex flex-1 flex-col items-center justify-center gap-3 p-5 text-center">
            <p className="text-sm text-muted-foreground">
              Ask about this file. The agent is instructed to read without changing project files.
            </p>
            <Button size="sm" variant="outline" onClick={onStart}>
              Open application agent
            </Button>
          </div>
        )}
      </aside>
    </div>
  );
}

export function ConsultancyFileViewer({
  environmentId,
  cwd,
  path,
  active = true,
  onAskAi,
  onAskSpreadsheet,
  projectId,
  onDirtyChange,
  spreadsheetChat,
  applicationChat,
  onStartApplicationAgent,
  excelAgentProfile,
  powerpointAgentProfile,
  spreadsheetThreadRef,
}: {
  environmentId: EnvironmentId;
  cwd: string;
  path: string;
  active?: boolean;
  onAskAi: (prompt: string) => boolean;
  projectId: ProjectId;
  onDirtyChange?: (dirty: boolean) => void;
  spreadsheetChat?: ReactNode;
  applicationChat?: ReactNode;
  onStartApplicationAgent: () => void;
  excelAgentProfile: AgentWorkspaceProfileSettings;
  powerpointAgentProfile: AgentWorkspaceProfileSettings;
  spreadsheetThreadRef?: ScopedThreadRef | undefined;
}) {
  const name = path.split("/").at(-1) ?? path;
  const isText = TEXT.test(path) && !SPREADSHEET.test(path);
  const presentationSourcePath = /\.pptx$/i.test(path) ? `${path.slice(0, -5)}.t3deck.json` : null;
  const presentationSource = useEnvironmentQuery(
    presentationSourcePath
      ? projectEnvironment.readFile({
          environmentId,
          input: { cwd, relativePath: presentationSourcePath },
        })
      : null,
  );
  const textQuery = useEnvironmentQuery(
    isText
      ? projectEnvironment.readFile({
          environmentId,
          input: { cwd, relativePath: path },
        })
      : null,
  );
  const resource = useMemo(
    () => ({ _tag: "draft-workspace-file" as const, cwd, path }),
    [cwd, path],
  );
  const asset = useAssetUrlState(environmentId, isText ? null : resource);
  const refreshAssetUrl = useAssetUrlRefresh(environmentId, isText ? null : resource);
  const assetUrl = asset._tag === "Success" ? asset.url : null;
  const [workbook, setWorkbook] = useState<File | null>(null);
  const workbookRef = useRef<File | null>(null);
  const wasActiveRef = useRef(active);
  const [loadError, setLoadError] = useState("");
  const [workbookDirty, setWorkbookDirty] = useState(false);
  const writeSpreadsheet = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const spreadsheetThread = useThread(spreadsheetThreadRef ?? null);
  const mutationId = latestWorkspaceMutationId(spreadsheetThread?.activities ?? []);

  useWorkspaceMutationRefresh({
    enabled: active && !workbookDirty && SPREADSHEET.test(path),
    mutationId,
    resourceKey: `spreadsheet:${environmentId}:${cwd}:${path}`,
    refresh: () => {
      void refreshAssetUrl()
        .then(async (url) => {
          if (!url) throw new Error("Could not refresh this workbook.");
          const response = await fetch(url, { cache: "no-store" });
          if (!response.ok) throw new Error("Could not refresh this workbook.");
          const file = new File([await response.blob()], name);
          workbookRef.current = file;
          setWorkbook(file);
          setLoadError("");
        })
        .catch((cause) =>
          setLoadError(cause instanceof Error ? cause.message : "Could not refresh this workbook."),
        );
    },
  });

  useEffect(() => {
    if (assetUrl === null || !SPREADSHEET.test(path)) return;
    const controller = new AbortController();
    void (async () => {
      try {
        const response = await fetch(assetUrl, { signal: controller.signal, cache: "no-store" });
        if (!response.ok) throw new Error("Could not open this file.");
        const file = new File([await response.blob()], name);
        if (controller.signal.aborted) return;
        workbookRef.current = file;
        setWorkbook(file);
      } catch (cause) {
        if (!controller.signal.aborted)
          setLoadError(cause instanceof Error ? cause.message : "Could not open this file.");
      }
    })();
    return () => controller.abort();
  }, [assetUrl, name, path]);

  useEffect(() => {
    const wasActive = wasActiveRef.current;
    wasActiveRef.current = active;
    if (wasActive || !active || !workbookRef.current || !SPREADSHEET.test(path)) return;
    const controller = new AbortController();
    void (async () => {
      try {
        const url = await refreshAssetUrl();
        if (!url || controller.signal.aborted) return;
        const response = await fetch(url, { signal: controller.signal, cache: "no-store" });
        if (!response.ok) throw new Error("Could not refresh this workbook.");
        const file = new File([await response.blob()], name);
        if (controller.signal.aborted) return;
        workbookRef.current = file;
        setWorkbook(file);
        setLoadError("");
      } catch (cause) {
        if (!controller.signal.aborted)
          setLoadError(cause instanceof Error ? cause.message : "Could not refresh this workbook.");
      }
    })();
    return () => controller.abort();
  }, [active, name, path, refreshAssetUrl]);

  if (workbook && SPREADSHEET.test(path))
    return (
      <div className="flex min-h-0 flex-1 flex-col">
        {loadError ? <p className="p-2 text-sm text-destructive">{loadError}</p> : null}
        <SpreadsheetFileViewer
          file={workbook}
          inline
          agentChat={spreadsheetChat}
          workspaceProfile={excelAgentProfile}
          onDirtyChange={setWorkbookDirty}
          onAskAi={(question, agentContext) => onAskSpreadsheet(question, agentContext, path)}
          sourcePath={`${cwd}/${path}`}
          onSave={async (data) => {
            const result = await writeSpreadsheet({
              environmentId,
              input: {
                cwd,
                relativePath: path,
                contents: workbookBase64(data),
                encoding: "base64",
              },
            });
            if (result._tag === "Failure") throw squashAtomCommandFailure(result);
          }}
        />
      </div>
    );

  if (isText) {
    if (textQuery.error) return <p className="p-5 text-sm text-destructive">{textQuery.error}</p>;
    if (!textQuery.data) return <p className="p-5 text-sm text-muted-foreground">Opening file…</p>;
    if (/\.t3deck\.json$/i.test(path))
      return (
        <div className="flex h-full min-h-0 flex-col">
          <PresentationDeckPreview
            contents={textQuery.data.contents}
            name={path}
            environmentId={environmentId}
            cwd={cwd}
            onAskAi={onAskAi}
            projectId={projectId}
            agentProfile={powerpointAgentProfile}
          />
        </div>
      );
    return (
      <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
        <div className="h-full overflow-auto p-5">
          {textQuery.data.truncated ? (
            <p className="mb-4 text-sm text-muted-foreground">
              Showing the beginning of this file.
            </p>
          ) : null}
          {MARKDOWN.test(path) ? (
            <ChatMarkdown text={textQuery.data.contents} cwd={cwd} />
          ) : (
            <pre className="whitespace-pre-wrap break-words font-sans text-sm leading-relaxed">
              {textQuery.data.contents}
            </pre>
          )}
        </div>
      </ApplicationFileLayout>
    );
  }

  if (presentationSourcePath && presentationSource.data)
    return (
      <div className="flex h-full min-h-0 flex-col">
        <PresentationDeckPreview
          contents={presentationSource.data.contents}
          name={presentationSourcePath}
          environmentId={environmentId}
          cwd={cwd}
          onAskAi={onAskAi}
          projectId={projectId}
          agentProfile={powerpointAgentProfile}
        />
      </div>
    );

  if (/\.docx$/i.test(path))
    return (
      <Suspense fallback={<p className="p-5 text-sm text-muted-foreground">Opening document…</p>}>
        <WordDocumentEditor
          environmentId={environmentId}
          cwd={cwd}
          path={path}
          projectId={projectId}
          assetUrl={assetUrl}
          assetError={asset._tag === "Failure"}
          onDirtyChange={onDirtyChange}
        />
      </Suspense>
    );

  if (asset._tag === "Failure" || loadError)
    return (
      <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
        <p className="p-5 text-sm text-destructive">{loadError || "Could not open this file."}</p>
      </ApplicationFileLayout>
    );
  if (asset._tag !== "Success")
    return (
      <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
        <p className="p-5 text-sm text-muted-foreground">Opening file…</p>
      </ApplicationFileLayout>
    );
  if (/\.pdf$/i.test(path))
    return (
      <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
        <BrowserDocumentFrame src={asset.url} title={name} pdf />
      </ApplicationFileLayout>
    );
  if (IMAGE.test(path))
    return (
      <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
        <img
          src={asset.url}
          alt={name}
          className="mx-auto max-h-full max-w-full object-contain p-5"
        />
      </ApplicationFileLayout>
    );
  if (SPREADSHEET.test(path))
    return <p className="p-5 text-sm text-muted-foreground">Opening workbook…</p>;
  return (
    <ApplicationFileLayout chat={applicationChat} onStart={onStartApplicationAgent}>
      <div className="p-5 text-sm">
        <p className="text-muted-foreground">Preview is unavailable for this file type.</p>
        <a className="mt-3 inline-block text-primary underline" href={asset.url} download={name}>
          Download file
        </a>
      </div>
    </ApplicationFileLayout>
  );
}

export function ConsultancyProjectFiles({
  environmentId,
  cwd,
  onOpenFile,
  onAskAi,
  projectId,
  powerpointAgentProfile,
}: {
  environmentId: EnvironmentId;
  cwd: string;
  onOpenFile: (path: string) => void;
  onAskAi: (prompt: string) => boolean;
  onAskSpreadsheet: (
    question: string,
    agentContext: string,
    path?: string,
  ) => boolean | Promise<boolean>;
  projectId: ProjectId;
  powerpointAgentProfile: AgentWorkspaceProfileSettings;
}) {
  const { entries, load, refresh, ready, error, isPending } = useDirectoryEntries(
    environmentId,
    cwd,
  );
  const [folder, setFolder] = useState("");
  const [newWordTitle, setNewWordTitle] = useState<string | null>(null);
  const [newExcelTitle, setNewExcelTitle] = useState<string | null>(null);
  const [wordError, setWordError] = useState("");
  const [excelError, setExcelError] = useState("");
  const [savingWord, setSavingWord] = useState(false);
  const [savingExcel, setSavingExcel] = useState(false);
  const standard = useEnvironmentSettings(
    environmentId,
    (settings) => settings.companyLibrary.wordStandard,
  );
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const listEntries = useAtomQueryRunner(projectEnvironment.listEntries, {
    reportFailure: false,
    refresh: true,
  });
  const visibleEntries = entries.filter((entry) => {
    const parent = entry.path.slice(0, Math.max(0, entry.path.lastIndexOf("/")));
    return parent === folder && !isConsultancyInternalPath(entry.path);
  });
  const createWordDocument = async () => {
    if (!newWordTitle?.trim()) return;
    setSavingWord(true);
    setWordError("");
    try {
      const path = wordDocumentPath(newWordTitle);
      const existing = await listEntries({ environmentId, input: { cwd, directoryPath: "word" } });
      if (existing._tag === "Failure") throw squashAtomCommandFailure(existing);
      if (existing.value.entries.some((entry) => entry.path.toLowerCase() === path.toLowerCase())) {
        throw new Error("A Word document with this name already exists.");
      }
      const contents = await exportWordDocument(
        newWordDocument(newWordTitle.trim().replace(/\.docx$/i, ""), "", standard),
        standard,
      );
      const saved = await writeFile({
        environmentId,
        input: { cwd, relativePath: path, contents, encoding: "base64" },
      });
      if (saved._tag === "Failure") throw squashAtomCommandFailure(saved);
      setNewWordTitle(null);
      setFolder("word");
      refresh();
      void load("word", true);
      onOpenFile(path);
    } catch (cause) {
      setWordError(cause instanceof Error ? cause.message : "Could not create Word document.");
    } finally {
      setSavingWord(false);
    }
  };
  const createExcelWorkbook = async () => {
    const title = newExcelTitle?.trim();
    if (!title) return;
    setSavingExcel(true);
    setExcelError("");
    try {
      const name = title.replace(/\.xlsx$/i, "");
      const safeName = name
        .replace(/[<>:"/\\|?*]/g, "-")
        .replace(/\p{Cc}/gu, "-")
        .slice(0, 90)
        .replace(/[. ]+$/, "");
      if (!safeName || safeName === "." || safeName === "..")
        throw new Error("Give the workbook a valid name.");
      const path = `excel/${safeName}.xlsx`;
      const existing = await listEntries({ environmentId, input: { cwd, directoryPath: "excel" } });
      if (existing._tag === "Failure") throw squashAtomCommandFailure(existing);
      if (existing.value.entries.some((entry) => entry.path.toLowerCase() === path.toLowerCase())) {
        throw new Error("A workbook with this name already exists.");
      }
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      workbook.addWorksheet("Sheet1");
      const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
      let binary = "";
      for (let offset = 0; offset < bytes.length; offset += 8192) {
        binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
      }
      const saved = await writeFile({
        environmentId,
        input: { cwd, relativePath: path, contents: btoa(binary), encoding: "base64" },
      });
      if (saved._tag === "Failure") throw squashAtomCommandFailure(saved);
      setNewExcelTitle(null);
      setFolder("excel");
      refresh();
      void load("excel", true);
      onOpenFile(path);
    } catch (cause) {
      setExcelError(cause instanceof Error ? cause.message : "Could not create the workbook.");
    } finally {
      setSavingExcel(false);
    }
  };

  return (
    <section className="mt-8">
      <div className="flex items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-medium">Project files</h2>
          <p className="mt-1 text-sm text-muted-foreground">
            Open a file in the project work area.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <Button variant="outline" size="sm" onClick={() => setNewExcelTitle("")}>
            <FileSpreadsheetIcon /> New Excel workbook
          </Button>
          <Button variant="outline" size="sm" onClick={() => setNewWordTitle("")}>
            New Word document
          </Button>
          <PresentationMaker
            environmentId={environmentId}
            cwd={cwd}
            buttonLabel="New presentation"
            onOpenInChat={onAskAi}
            projectId={projectId}
            agentProfile={powerpointAgentProfile}
            onSaved={(path) => {
              refresh();
              setFolder("powerpoints");
              void load("powerpoints");
              onOpenFile(path);
            }}
          />
          <Button variant="ghost" size="sm" onClick={refresh} disabled={isPending}>
            Refresh
          </Button>
        </div>
      </div>
      {newExcelTitle !== null ? (
        <div className="mt-3 flex flex-wrap items-center gap-2 rounded-md border border-border p-3">
          <Input
            aria-label="New Excel workbook name"
            placeholder="Workbook name"
            value={newExcelTitle}
            onChange={(event) => setNewExcelTitle(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") void createExcelWorkbook();
            }}
          />
          <Button
            size="sm"
            disabled={!newExcelTitle.trim() || savingExcel}
            onClick={() => void createExcelWorkbook()}
          >
            {savingExcel ? "Creating..." : "Create"}
          </Button>
          <Button
            variant="ghost"
            size="sm"
            onClick={() => {
              setNewExcelTitle(null);
              setExcelError("");
            }}
          >
            Cancel
          </Button>
          {excelError ? (
            <p role="alert" className="w-full text-xs text-destructive">
              {excelError}
            </p>
          ) : null}
        </div>
      ) : null}
      {newWordTitle !== null ? (
        <div className="mt-3 flex flex-wrap items-center gap-2 rounded-md border border-border p-3">
          <Input
            aria-label="New Word document name"
            placeholder="Document name"
            value={newWordTitle}
            onChange={(event) => setNewWordTitle(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") void createWordDocument();
            }}
          />
          <Button
            size="sm"
            disabled={!newWordTitle.trim() || savingWord}
            onClick={() => void createWordDocument()}
          >
            {savingWord ? "Creating…" : "Create"}
          </Button>
          <Button
            variant="ghost"
            size="sm"
            onClick={() => {
              setNewWordTitle(null);
              setWordError("");
            }}
          >
            Cancel
          </Button>
          {wordError ? (
            <p role="alert" className="w-full text-xs text-destructive">
              {wordError}
            </p>
          ) : null}
        </div>
      ) : null}
      <div className="mt-4 overflow-hidden rounded-xl border border-border bg-card">
        {folder ? (
          <button
            type="button"
            className="flex w-full items-center gap-2 border-b border-border px-4 py-3 text-left text-sm hover:bg-muted/50"
            onClick={() => {
              setFolder(folder.slice(0, Math.max(0, folder.lastIndexOf("/"))));
            }}
          >
            <ChevronLeftIcon className="size-4" /> Back
          </button>
        ) : null}
        {visibleEntries.map((entry) => (
          <button
            key={entry.path}
            type="button"
            className="flex w-full items-center gap-3 border-b border-border px-4 py-3 text-left text-sm last:border-b-0 hover:bg-muted/50"
            onClick={() => {
              if (entry.kind === "directory") {
                setFolder(entry.path);
                void load(entry.path);
              } else {
                onOpenFile(entry.path);
              }
            }}
          >
            {entry.kind === "directory" ? (
              <FolderIcon className="size-4 shrink-0 text-primary" />
            ) : SPREADSHEET.test(entry.path) ? (
              <FileSpreadsheetIcon className="size-4 shrink-0 text-primary" />
            ) : (
              <FileTextIcon className="size-4 shrink-0 text-primary" />
            )}
            <span className="min-w-0 flex-1 truncate">{entry.path.split("/").at(-1)}</span>
          </button>
        ))}
        {error ? <p className="p-5 text-sm text-destructive">{error}</p> : null}
        {ready && !isPending && !error && visibleEntries.length === 0 ? (
          <p className="p-5 text-sm text-muted-foreground">No files in this folder.</p>
        ) : null}
        {!ready && !error ? (
          <p className="p-5 text-sm text-muted-foreground">Loading files…</p>
        ) : null}
      </div>
    </section>
  );
}
