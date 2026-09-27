import { useAtomValue } from "@effect/atom-react";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";
import {
  DEFAULT_COMPANY_WORD_STANDARD,
  type EnvironmentId,
  type ProjectId,
} from "@t3tools/contracts";
import { EditorContent, useEditor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import {
  BoldIcon,
  DownloadIcon,
  ItalicIcon,
  MessageSquareIcon,
  RefreshCwIcon,
  SaveIcon,
} from "lucide-react";
import { useEffect, useMemo, useState, type CSSProperties } from "react";

import { useAssetUrlRefresh } from "~/assets/assetUrls";
import { ThreadRouteView } from "~/components/ThreadRouteView";
import { Button } from "~/components/ui/button";
import {
  latestWorkspaceMutationId,
  useWorkspaceMutationRefresh,
} from "~/hooks/useWorkspaceMutationRefresh";
import { useThread } from "~/state/entities";
import { projectEnvironment } from "~/state/projects";
import { serverEnvironment } from "~/state/server";
import { useAtomCommand } from "~/state/use-atom-command";
import { wordAgentContext } from "./editorAgentContext";
import { useFileChatThread } from "./useFileChatThread";
import {
  applyWordStructure,
  exportWordDocument,
  openWordDocument,
  type OpenedWordDocument,
} from "./wordDocument";
import "./WordDocumentEditor.css";

interface WordDocumentEditorProps {
  readonly environmentId: EnvironmentId;
  readonly cwd: string;
  readonly path: string;
  readonly projectId: ProjectId;
  readonly assetUrl: string | null;
  readonly assetError: boolean;
  readonly onDirtyChange?: ((dirty: boolean) => void) | undefined;
}

async function readWordFile(url: string, signal?: AbortSignal) {
  const response = await fetch(url, { ...(signal ? { signal } : {}), cache: "no-store" });
  if (!response.ok) throw new Error("Could not open the Word document.");
  const buffer = await response.arrayBuffer();
  return { document: await openWordDocument(buffer), bytes: new Uint8Array(buffer) };
}

function WordEditorCanvas({
  document,
  sourceBytes,
  assetUrl,
  environmentId,
  cwd,
  path,
  projectId,
  onReload,
  onSavedBytes,
  checkForExternalChanges,
  onDirtyChange,
}: {
  document: OpenedWordDocument;
  sourceBytes: Uint8Array;
  assetUrl: string | null;
  environmentId: EnvironmentId;
  cwd: string;
  path: string;
  projectId: ProjectId;
  onReload: () => Promise<void>;
  onSavedBytes: (bytes: Uint8Array) => void;
  checkForExternalChanges: (sourceBytes: Uint8Array) => Promise<void>;
  onDirtyChange?: ((dirty: boolean) => void) | undefined;
}) {
  const standard =
    useAtomValue(serverEnvironment.settingsValueAtom(environmentId))?.companyLibrary.wordStandard ??
    DEFAULT_COMPANY_WORD_STANDARD;
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const {
    ensureThread,
    target: chatTarget,
    threadRef: activeRef,
  } = useFileChatThread(environmentId, projectId, cwd, path);
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [chatSelection, setChatSelection] = useState("");
  const chatThread = useThread(activeRef);
  const mutationId = latestWorkspaceMutationId(chatThread?.activities ?? []);
  useWorkspaceMutationRefresh({
    enabled: !dirty && !saving,
    mutationId,
    refresh: () => {
      void onReload();
    },
    resourceKey: `word:${environmentId}:${cwd}:${path}`,
  });
  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: { levels: [1, 2] },
        blockquote: false,
        bulletList: false,
        orderedList: false,
        listItem: false,
        codeBlock: false,
        code: false,
        horizontalRule: false,
        link: false,
        underline: false,
        strike: false,
        dropcursor: false,
        gapcursor: false,
      }),
    ],
    content: document.content,
    editable: document.editable,
    immediatelyRender: false,
    onUpdate: () => setDirty(true),
  });

  useEffect(() => {
    if (!editor) return;
    editor.commands.setContent(document.content, { emitUpdate: false });
    editor.setEditable(document.editable);
  }, [document, editor]);

  useEffect(() => {
    if (!dirty) return;
    const warn = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);

  useEffect(() => onDirtyChange?.(dirty), [dirty, onDirtyChange]);

  const name = path.split("/").at(-1) ?? "Document.docx";
  const download = async () => {
    if (!editor || !document.editable) return;
    setError("");
    try {
      const base64 = await exportWordDocument(editor.getJSON(), standard);
      const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
      const url = URL.createObjectURL(
        new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
      );
      const link = window.document.createElement("a");
      link.href = url;
      link.download = name;
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not export document.");
    }
  };
  const save = async () => {
    if (!editor || !document.editable) return;
    setSaving(true);
    setError("");
    try {
      await checkForExternalChanges(sourceBytes);
      const snapshot = editor.getJSON();
      const contents = await exportWordDocument(snapshot, standard);
      const result = await writeFile({
        environmentId,
        input: { cwd, relativePath: path, contents, encoding: "base64" },
      });
      if (result._tag === "Failure") throw squashAtomCommandFailure(result);
      onSavedBytes(Uint8Array.from(atob(contents), (character) => character.charCodeAt(0)));
      if (JSON.stringify(editor.getJSON()) === JSON.stringify(snapshot)) setDirty(false);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not save document.");
    } finally {
      setSaving(false);
    }
  };
  const ask = async (selectionOnly: boolean) => {
    if (!editor) return;
    setError("");
    try {
      await ensureThread();
      const { from, to } = editor.state.selection;
      const selected =
        selectionOnly && from !== to
          ? editor.state.doc.textBetween(from, to, " ").slice(0, 2_000)
          : "";
      setChatSelection(selected);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not start the document chat.");
    }
  };
  const reload = () => {
    if (dirty && !window.confirm("Discard unsaved document changes and reload from disk?")) return;
    void onReload()
      .then(() => setDirty(false))
      .catch((cause) =>
        setError(cause instanceof Error ? cause.message : "Could not reload document."),
      );
  };
  const themeStyle = {
    fontFamily: standard.fontFamily,
    color: `#${standard.bodyColor}`,
    "--word-heading": `#${standard.headingColor}`,
    "--word-accent": `#${standard.accentColor}`,
  } as CSSProperties;

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex flex-wrap items-center gap-2 border-b border-border px-4 py-2">
        <span className="mr-auto min-w-0 truncate text-sm font-medium">
          {name}
          {dirty ? " · Unsaved changes" : ""}
        </span>
        {document.editable && editor ? (
          <>
            <Button
              size="sm"
              variant="outline"
              onClick={() => editor.chain().focus().toggleBold().run()}
              aria-label="Bold"
            >
              <BoldIcon className="size-4" />
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => editor.chain().focus().toggleItalic().run()}
              aria-label="Italic"
            >
              <ItalicIcon className="size-4" />
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => editor.chain().focus().setParagraph().run()}
            >
              Text
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => editor.chain().focus().toggleHeading({ level: 1 }).run()}
            >
              Title
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => editor.chain().focus().toggleHeading({ level: 2 }).run()}
            >
              Heading
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => {
                editor.commands.setContent(applyWordStructure(editor.getJSON(), standard));
                setDirty(true);
              }}
            >
              Apply company standard
            </Button>
            <Button size="sm" variant="outline" onClick={() => void download()}>
              <DownloadIcon className="size-4" /> Download
            </Button>
            <Button size="sm" disabled={!dirty || saving} onClick={() => void save()}>
              <SaveIcon className="size-4" /> {saving ? "Saving…" : "Save"}
            </Button>
          </>
        ) : null}
        {!document.editable && assetUrl ? (
          <a className="text-sm text-primary underline" href={assetUrl} download={name}>
            Download original
          </a>
        ) : null}
        <Button size="sm" variant="outline" onClick={reload}>
          <RefreshCwIcon className="size-4" /> Reload
        </Button>
        <Button size="sm" variant="outline" onClick={() => void ask(false)}>
          <MessageSquareIcon className="size-4" /> Chat
        </Button>
      </div>
      {document.reason ? (
        <p className="border-b border-border bg-muted/40 px-4 py-2 text-xs text-muted-foreground">
          {document.reason}
        </p>
      ) : null}
      {error ? (
        <p
          role="alert"
          className="border-b border-destructive/30 px-4 py-2 text-xs text-destructive"
        >
          {error}
        </p>
      ) : null}
      <div className="flex min-h-0 flex-1 flex-col lg:flex-row">
        <div className="min-h-0 flex-1 overflow-auto bg-muted/30 p-4 sm:p-8">
          <div
            data-word-document-page=""
            className="mx-auto min-h-[70rem] max-w-[52rem] bg-white px-8 py-10 shadow-sm sm:px-14 sm:py-16"
            style={themeStyle}
          >
            {editor ? <EditorContent editor={editor} /> : null}
            {document.editable && editor ? (
              <Button size="sm" variant="ghost" onClick={() => void ask(true)}>
                Ask about selection
              </Button>
            ) : null}
          </div>
        </div>
        {chatTarget ? (
          <aside className="flex min-h-64 flex-col border-t border-border lg:w-[min(42%,34rem)] lg:border-l lg:border-t-0">
            <ThreadRouteView
              target={chatTarget}
              embedded
              editorContext={wordAgentContext(path, document.editable, chatSelection)}
            />
          </aside>
        ) : null}
      </div>
    </div>
  );
}

export function WordDocumentEditor(props: WordDocumentEditorProps) {
  const resource = useMemo(
    () => ({ _tag: "draft-workspace-file" as const, cwd: props.cwd, path: props.path }),
    [props.cwd, props.path],
  );
  const refreshAssetUrl = useAssetUrlRefresh(props.environmentId, resource);
  const [opened, setOpened] = useState<{
    document: OpenedWordDocument;
    bytes: Uint8Array;
  } | null>(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!props.assetUrl || opened) return;
    const controller = new AbortController();
    void readWordFile(props.assetUrl, controller.signal)
      .then((next) => {
        if (!controller.signal.aborted) setOpened(next);
      })
      .catch((cause) => {
        if (!controller.signal.aborted)
          setError(cause instanceof Error ? cause.message : "Could not open the Word document.");
      });
    return () => controller.abort();
  }, [opened, props.assetUrl]);
  const reload = async () => {
    setLoading(true);
    setError("");
    try {
      const url = (await refreshAssetUrl()) ?? props.assetUrl;
      if (!url) throw new Error("Could not locate the Word document.");
      const next = await readWordFile(url);
      setOpened(next);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not reload the Word document.");
    } finally {
      setLoading(false);
    }
  };
  const checkForExternalChanges = async (sourceBytes: Uint8Array) => {
    const url = (await refreshAssetUrl()) ?? props.assetUrl;
    if (!url) throw new Error("Could not check the current Word document.");
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error("Could not check the current Word document.");
    const current = new Uint8Array(await response.arrayBuffer());
    if (
      current.length !== sourceBytes.length ||
      current.some((byte, index) => byte !== sourceBytes[index])
    ) {
      throw new Error("This document changed on disk. Reload it before saving your edits.");
    }
  };

  if (!opened && (props.assetError || error))
    return (
      <p role="alert" className="p-5 text-sm text-destructive">
        {error || "Could not open the Word document."}
      </p>
    );
  if (!opened) return <p className="p-5 text-sm text-muted-foreground">Opening document…</p>;
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      {error ? (
        <p role="alert" className="border-b px-4 py-1 text-xs text-destructive">
          {error}
        </p>
      ) : null}
      {loading ? (
        <p className="border-b px-4 py-1 text-xs text-muted-foreground">Reloading document…</p>
      ) : null}
      <WordEditorCanvas
        document={opened.document}
        sourceBytes={opened.bytes}
        assetUrl={props.assetUrl}
        environmentId={props.environmentId}
        cwd={props.cwd}
        path={props.path}
        projectId={props.projectId}
        onReload={reload}
        onSavedBytes={(bytes) =>
          setOpened((current) => (current ? { ...current, bytes } : current))
        }
        checkForExternalChanges={checkForExternalChanges}
        onDirtyChange={props.onDirtyChange}
      />
    </div>
  );
}
