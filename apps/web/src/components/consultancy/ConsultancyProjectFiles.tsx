import type { EnvironmentId } from "@t3tools/contracts";
import { ChevronLeftIcon, FolderIcon, FileSpreadsheetIcon, FileTextIcon } from "lucide-react";
import { useEffect, useMemo, useState } from "react";

import { useAssetUrlState } from "~/assets/assetUrls";
import ChatMarkdown from "~/components/ChatMarkdown";
import { BrowserDocumentFrame } from "~/components/files/BrowserDocumentFrame";
import { useDirectoryEntries } from "~/components/files/useDirectoryEntries";
import { Button } from "~/components/ui/button";
import { projectEnvironment } from "~/state/projects";
import { useEnvironmentQuery } from "~/state/query";

import { extractDocumentContext } from "./documentContext";
import { isConsultancyInternalPath } from "./consultancyData";

const SPREADSHEET = /\.(xlsx|csv|tsv)$/i;
const MARKDOWN = /\.md$/i;
const TEXT = /\.(md|txt|json|xml|html|csv|tsv)$/i;
const IMAGE = /\.(png|jpe?g|gif|webp|svg)$/i;

function FilePreview({
  environmentId,
  cwd,
  path,
  onOpenWorkbook,
}: {
  environmentId: EnvironmentId;
  cwd: string;
  path: string;
  onOpenWorkbook: (file: File) => void;
}) {
  const name = path.split("/").at(-1) ?? path;
  const isText = TEXT.test(path) && !SPREADSHEET.test(path);
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
  const assetUrl = asset._tag === "Success" ? asset.url : null;
  const [document, setDocument] = useState<{ path: string; text: string } | null>(null);
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    if (assetUrl === null || (!SPREADSHEET.test(path) && !/\.docx$/i.test(path))) return;
    const controller = new AbortController();
    void (async () => {
      try {
        const response = await fetch(assetUrl, { signal: controller.signal });
        if (!response.ok) throw new Error("Could not open this file.");
        const file = new File([await response.blob()], name);
        if (controller.signal.aborted) return;
        if (SPREADSHEET.test(path)) {
          onOpenWorkbook(file);
        } else {
          const extracted = await extractDocumentContext(file);
          if (!controller.signal.aborted) setDocument({ path, text: extracted.text });
        }
      } catch (cause) {
        if (!controller.signal.aborted)
          setLoadError(cause instanceof Error ? cause.message : "Could not open this file.");
      }
    })();
    return () => controller.abort();
  }, [assetUrl, name, onOpenWorkbook, path]);

  if (isText) {
    if (textQuery.error) return <p className="p-5 text-sm text-destructive">{textQuery.error}</p>;
    if (!textQuery.data) return <p className="p-5 text-sm text-muted-foreground">Opening file…</p>;
    return (
      <div className="max-h-[40rem] overflow-auto p-5">
        {textQuery.data.truncated ? (
          <p className="mb-4 text-sm text-muted-foreground">Showing the beginning of this file.</p>
        ) : null}
        {MARKDOWN.test(path) ? (
          <ChatMarkdown text={textQuery.data.contents} cwd={cwd} />
        ) : (
          <pre className="whitespace-pre-wrap break-words font-sans text-sm leading-relaxed">
            {textQuery.data.contents}
          </pre>
        )}
      </div>
    );
  }

  if (asset._tag === "Failure" || loadError)
    return (
      <p className="p-5 text-sm text-destructive">{loadError || "Could not open this file."}</p>
    );
  if (asset._tag !== "Success")
    return <p className="p-5 text-sm text-muted-foreground">Opening file…</p>;
  if (/\.pdf$/i.test(path))
    return (
      <div className="flex h-[40rem] min-h-0 flex-col">
        <BrowserDocumentFrame src={asset.url} title={name} pdf />
      </div>
    );
  if (IMAGE.test(path))
    return <img src={asset.url} alt={name} className="mx-auto max-h-[40rem] max-w-full p-5" />;
  if (/\.docx$/i.test(path))
    return document?.path === path ? (
      <pre className="max-h-[40rem] overflow-auto whitespace-pre-wrap break-words p-5 font-sans text-sm leading-relaxed">
        {document.text}
      </pre>
    ) : (
      <p className="p-5 text-sm text-muted-foreground">Opening document…</p>
    );
  if (SPREADSHEET.test(path))
    return <p className="p-5 text-sm text-muted-foreground">Opening workbook…</p>;
  return (
    <div className="p-5 text-sm">
      <p className="text-muted-foreground">Preview is unavailable for this file type.</p>
      <a className="mt-3 inline-block text-primary underline" href={asset.url} download={name}>
        Download file
      </a>
    </div>
  );
}

export function ConsultancyProjectFiles({
  environmentId,
  cwd,
  onOpenWorkbook,
}: {
  environmentId: EnvironmentId;
  cwd: string;
  onOpenWorkbook: (file: File) => void;
}) {
  const { entries, load, refresh, ready, error, isPending } = useDirectoryEntries(
    environmentId,
    cwd,
  );
  const [folder, setFolder] = useState("");
  const [selectedPath, setSelectedPath] = useState<string | null>(null);
  const visibleEntries = entries.filter((entry) => {
    const parent = entry.path.slice(0, Math.max(0, entry.path.lastIndexOf("/")));
    return parent === folder && !isConsultancyInternalPath(entry.path);
  });

  return (
    <section className="mt-8">
      <div className="flex items-center justify-between gap-3">
        <div>
          <h2 className="text-lg font-medium">Project files</h2>
          <p className="mt-1 text-sm text-muted-foreground">Open a file to read it here.</p>
        </div>
        <Button variant="ghost" size="sm" onClick={refresh} disabled={isPending}>
          Refresh
        </Button>
      </div>
      <div className="mt-4 overflow-hidden rounded-xl border border-border bg-card">
        {folder ? (
          <button
            type="button"
            className="flex w-full items-center gap-2 border-b border-border px-4 py-3 text-left text-sm hover:bg-muted/50"
            onClick={() => {
              setFolder(folder.slice(0, Math.max(0, folder.lastIndexOf("/"))));
              setSelectedPath(null);
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
            aria-pressed={entry.kind === "file" && selectedPath === entry.path}
            onClick={() => {
              if (entry.kind === "directory") {
                setFolder(entry.path);
                setSelectedPath(null);
                void load(entry.path);
              } else {
                setSelectedPath(entry.path);
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
      {selectedPath ? (
        <div className="mt-4 overflow-hidden rounded-xl border border-border bg-card">
          <div className="border-b border-border px-5 py-3 text-sm font-medium">
            {selectedPath.split("/").at(-1)}
          </div>
          <FilePreview
            key={selectedPath}
            environmentId={environmentId}
            cwd={cwd}
            path={selectedPath}
            onOpenWorkbook={onOpenWorkbook}
          />
        </div>
      ) : null}
    </section>
  );
}
