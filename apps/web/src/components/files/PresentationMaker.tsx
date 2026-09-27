import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";
import { scopeProjectRef, scopeThreadRef } from "@t3tools/client-runtime/environment";
import {
  ArrowDownIcon,
  ArrowUpIcon,
  DownloadIcon,
  PlusIcon,
  PresentationIcon,
  MousePointer2Icon,
  SparklesIcon,
  Trash2Icon,
} from "lucide-react";
import type { EnvironmentId, ProjectId, ScopedThreadRef, ThreadId } from "@t3tools/contracts";
import { CONSULTANCY_PRESENTATION_STANDARD } from "@t3tools/shared/consultancyPresentationStandard";
import { CONSULTANCY_CHART_TYPES } from "@t3tools/shared/consultancyChart";
import { ConsultancyChartView } from "~/components/charts/ConsultancyChartView";
import { type DraftId, useComposerDraftStore } from "~/composerDraftStore";
import { useNewThreadHandler } from "~/hooks/useHandleNewThread";
import { useThread, useThreadShell } from "~/state/entities";
import {
  latestWorkspaceMutationId,
  useWorkspaceMutationRefresh,
} from "~/hooks/useWorkspaceMutationRefresh";
import { projectEnvironment } from "~/state/projects";
import { useAtomCommand } from "~/state/use-atom-command";
import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import { Dialog, DialogPopup, DialogTitle } from "~/components/ui/dialog";
import {
  confirmProjectFileQueryData,
  setProjectFileQueryData,
  useProjectFileQuery,
} from "./projectFilesQueryState";
import { downloadPresentation } from "./presentationExport";
import { presentationAgentContext } from "./editorAgentContext";
import { buildPresentation, importPresentationChart } from "./presentationWork";
import {
  deckPath,
  newDeck,
  newSlide,
  parseDeck,
  type PresentationChart,
  type PresentationDeck,
  type PresentationSlide,
  type SlideLayout,
} from "./presentationDeck";

const LAYOUTS: { value: SlideLayout; label: string }[] = [
  { value: "cover", label: "Cover" },
  { value: "section", label: "Section" },
  { value: "content", label: "Key points" },
  { value: "two-column", label: "Two columns" },
  { value: "chart", label: "Chart" },
];

const EmbeddedThread = lazy(() =>
  import("~/components/ThreadRouteView").then((module) => ({ default: module.ThreadRouteView })),
);

function PresentationChat({
  environmentId,
  projectId,
  threadRef,
  editorContext,
  onThreadRefChange,
  onBeforeStart,
  onThreadCreated,
}: {
  environmentId: EnvironmentId;
  projectId: ProjectId;
  threadRef?: ScopedThreadRef;
  editorContext: string;
  onThreadRefChange: (ref: ScopedThreadRef | null) => void;
  onBeforeStart: () => Promise<string>;
  onThreadCreated: (path: string, threadId: ThreadId) => Promise<void>;
}) {
  const newThread = useNewThreadHandler();
  const [draftId, setDraftId] = useState<DraftId | null>(null);
  const [reservedRef, setReservedRef] = useState<ScopedThreadRef | null>(null);
  const [error, setError] = useState("");
  const [pending, setPending] = useState(false);
  const draftSession = useComposerDraftStore((store) =>
    draftId ? store.getDraftSession(draftId) : null,
  );
  const reservedShell = useThreadShell(reservedRef);
  const activeRef = draftId
    ? (draftSession?.promotedTo ?? (reservedShell ? reservedRef : null))
    : (threadRef ?? null);
  useEffect(() => onThreadRefChange(activeRef), [activeRef, onThreadRefChange]);
  const target = activeRef
    ? { kind: "server" as const, threadRef: activeRef }
    : draftId
      ? { kind: "draft" as const, draftId }
      : null;
  const start = async () => {
    setPending(true);
    setError("");
    try {
      const path = await onBeforeStart();
      const result = await newThread(scopeProjectRef(environmentId, projectId), {
        navigate: false,
      });
      if (!result) throw new Error("Could not start the presentation chat.");
      setDraftId(result.draftId);
      setReservedRef(scopeThreadRef(environmentId, result.threadId));
      await onThreadCreated(path, result.threadId);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not start the presentation chat.");
    } finally {
      setPending(false);
    }
  };

  if (!target)
    return (
      <div className="flex flex-1 flex-col items-center justify-center gap-3 p-5 text-center">
        <p className="text-sm text-muted-foreground">Ask about this deck or describe a change.</p>
        <Button size="sm" disabled={pending} onClick={() => void start()}>
          {pending ? "Opening chat…" : "Start chat"}
        </Button>
        {error ? (
          <p role="alert" className="text-xs text-destructive">
            {error}
          </p>
        ) : null}
      </div>
    );
  return (
    <div className="min-h-0 flex-1 overflow-hidden">
      <Suspense fallback={<p className="p-4 text-sm text-muted-foreground">Opening chat…</p>}>
        <EmbeddedThread target={target} embedded editorContext={editorContext} />
      </Suspense>
    </div>
  );
}

function SlideCanvas({
  slide,
  index,
  selecting,
  selectedElement,
  onSelectElement,
}: {
  slide: PresentationSlide;
  index: number;
  selecting: boolean;
  selectedElement: string | null;
  onSelectElement: (element: string) => void;
}) {
  const { style, colors } = CONSULTANCY_PRESENTATION_STANDARD;
  const headline = slide.layout === "cover" || slide.layout === "section";
  return (
    <div
      className={`relative aspect-video overflow-hidden rounded-lg p-8 shadow-sm ${headline ? "border-l-4" : "border-t-4"}`}
      style={{
        background: `#${headline ? colors.navy : style.background}`,
        color: `#${headline ? style.background : style.foreground}`,
        borderColor: `#${headline ? colors.gold : colors.navy}`,
        fontFamily: style.fontFace,
      }}
    >
      {!headline ? (
        <span className="absolute left-6 top-7 h-5 w-1" style={{ background: `#${colors.gold}` }} />
      ) : null}
      <p
        className={`text-3xs font-semibold tracking-widest ${headline ? "text-white/75" : "ml-3"}`}
        style={headline ? undefined : { color: `#${style.muted}` }}
      >
        BERENSCHOT 2026 / PRESENTATIE
      </p>
      <h2
        className={`${headline ? "mt-12 text-4xl font-semibold" : "mt-6 text-3xl font-semibold"} ${selecting ? "cursor-pointer hover:outline hover:outline-2 hover:outline-offset-4 hover:outline-primary" : ""} ${selectedElement === "title" ? "outline outline-2 outline-offset-4 outline-primary" : ""}`}
        onClick={() => selecting && onSelectElement("title")}
      >
        {slide.title}
      </h2>
      <div className="mt-5 h-1 w-16" style={{ background: `#${colors.gold}` }} />
      {slide.layout === "chart" ? (
        <div
          className={`mt-6 max-h-[52%] overflow-auto ${selecting ? "cursor-pointer hover:outline hover:outline-2 hover:outline-primary" : ""} ${selectedElement === "chart" ? "outline outline-2 outline-primary" : ""}`}
          onClick={() => selecting && onSelectElement("chart")}
        >
          {slide.chart ? (
            <ConsultancyChartView chart={slide.chart} compact />
          ) : (
            <p className="text-sm opacity-70">Importeer data of kies een grafiek.</p>
          )}
        </div>
      ) : slide.layout === "two-column" ? (
        <div className="mt-7 grid max-h-[56%] grid-cols-2 gap-4 overflow-auto text-lg">
          <p
            className={`rounded-lg p-5 whitespace-pre-line ${selecting ? "cursor-pointer hover:outline hover:outline-2 hover:outline-offset-2 hover:outline-primary" : ""} ${selectedElement === "body" ? "outline outline-2 outline-primary" : ""}`}
            style={{ background: `#${colors.paleBlue}` }}
            onClick={() => selecting && onSelectElement("body")}
          >
            <span
              className="mb-3 block text-xs font-semibold tracking-wider"
              style={{ color: `#${style.accent}` }}
            >
              INZICHT
            </span>
            {slide.body}
          </p>
          <p
            className="rounded-lg border p-5 whitespace-pre-line"
            style={{ borderColor: "#D8E1EA" }}
            onClick={() => selecting && onSelectElement("right column")}
          >
            <span
              className="mb-3 block text-xs font-semibold tracking-wider"
              style={{ color: `#${colors.warmBrown}` }}
            >
              IMPLICATIE
            </span>
            {slide.rightBody}
          </p>
        </div>
      ) : (
        <p
          className="mt-7 max-h-[58%] overflow-auto whitespace-pre-line text-lg leading-relaxed"
          style={{ color: headline ? "#DCE7F2" : undefined }}
          onClick={() => selecting && onSelectElement("body")}
        >
          {slide.body}
        </p>
      )}
      <span
        className="absolute bottom-4 right-6 text-xs"
        style={{ color: headline ? "#C7D7EA" : `#${style.muted}` }}
      >
        {index + 1}
      </span>
      {slide.source || slide.chart?.source ? (
        <span
          className="absolute bottom-4 left-6 max-w-[75%] truncate text-3xs"
          style={{ color: headline ? "#C7D7EA" : `#${style.muted}` }}
        >
          Bron: {slide.source ?? slide.chart?.source}
        </span>
      ) : null}
    </div>
  );
}

interface EditorProps {
  environmentId: EnvironmentId;
  cwd: string;
  initialDeck: PresentationDeck;
  relativePath?: string;
  onSaved?: (path: string) => void;
  onAskAi?: (prompt: string) => boolean;
  projectId?: ProjectId;
}

function PresentationEditor({
  environmentId,
  cwd,
  initialDeck,
  relativePath,
  onSaved,
  onAskAi,
  projectId,
}: EditorProps) {
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const [deck, setDeck] = useState(initialDeck);
  const [dirty, setDirty] = useState(false);
  const [selecting, setSelecting] = useState(false);
  const [selectedElement, setSelectedElement] = useState<string | null>(null);
  const [history, setHistory] = useState<Array<{ role: "user" | "status"; text: string }>>([]);
  const [selected, setSelected] = useState(0);
  const [brief, setBrief] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [savedPath, setSavedPath] = useState(relativePath);
  const pendingLocalSave = useRef<string | null>(null);
  const lastExportedSource = useRef<string | null>(null);
  const [chatThreadRef, setChatThreadRef] = useState<ScopedThreadRef | null>(
    deck.chatThreadId ? scopeThreadRef(environmentId, deck.chatThreadId) : null,
  );
  const chatThread = useThread(chatThreadRef);
  const sourceFile = useProjectFileQuery(
    environmentId,
    cwd,
    savedPath ?? null,
    savedPath !== undefined,
  );
  const mutationId = useMemo(
    () => latestWorkspaceMutationId(chatThread?.activities ?? []),
    [chatThread?.activities],
  );
  useWorkspaceMutationRefresh({
    enabled: savedPath !== undefined,
    mutationId,
    refresh: sourceFile.refresh,
    resourceKey: `presentation:${environmentId}:${cwd}:${savedPath ?? ""}`,
  });
  useEffect(() => {
    if (dirty || !sourceFile.data?.contents) return;
    if (pendingLocalSave.current !== null) {
      if (sourceFile.data.contents !== pendingLocalSave.current) return;
      pendingLocalSave.current = null;
    }
    try {
      const nextDeck = parseDeck(sourceFile.data.contents);
      if (JSON.stringify(nextDeck) === JSON.stringify(deck)) return;
      setDeck(nextDeck);
      if (lastExportedSource.current !== sourceFile.data.contents) {
        lastExportedSource.current = sourceFile.data.contents;
        const pptxPath = (savedPath ?? "").replace(/\.t3deck\.json$/i, ".pptx");
        void buildPresentation(nextDeck)
          .then((contents) =>
            writeFile({
              environmentId,
              input: { cwd, relativePath: pptxPath, contents, encoding: "base64" },
            }),
          )
          .then((result) => {
            if (result._tag === "Failure") throw new Error("Could not update PowerPoint export.");
          })
          .catch((cause) =>
            setError(
              cause instanceof Error ? cause.message : "Could not update PowerPoint export.",
            ),
          );
      }
    } catch {
      // A partial agent write is replaced by the next completed file update.
    }
  }, [cwd, deck, dirty, environmentId, savedPath, sourceFile.data?.contents, writeFile]);
  const lastDeckFromFile = useRef(initialDeck);
  const selectedSlide = deck.slides[selected];

  useEffect(() => {
    if (initialDeck === lastDeckFromFile.current) return;
    lastDeckFromFile.current = initialDeck;
    if (dirty) return;
    setDeck(initialDeck);
  }, [initialDeck, dirty]);

  const editDeck = (change: Partial<PresentationDeck>) => {
    setDirty(true);
    setDeck((current) => ({ ...current, ...change }));
  };
  const editSlide = (change: Partial<PresentationSlide>) => {
    setDirty(true);
    setDeck((current) => ({
      ...current,
      slides: current.slides.map((slide, index) =>
        index === selected ? { ...slide, ...change } : slide,
      ),
    }));
  };
  const save = async (download: boolean): Promise<string> => {
    if (!deck.title.trim()) throw new Error("Give this presentation a title.");
    if (!deck.slides.length) throw new Error("Add at least one slide.");
    setBusy(true);
    setError("");
    try {
      const base = savedPath?.replace(/\.t3deck\.json$/, "") ?? deckPath(deck.title);
      const sourcePath = `${base}.t3deck.json`;
      const pptxPath = `${base}.pptx`;
      const contents = JSON.stringify(deck, null, 2);
      const base64 = await buildPresentation(deck);
      const sourceResult = await writeFile({
        environmentId,
        input: { cwd, relativePath: sourcePath, contents },
      });
      if (sourceResult._tag === "Failure") throw new Error("Could not save the editable deck.");
      const pptxResult = await writeFile({
        environmentId,
        input: { cwd, relativePath: pptxPath, contents: base64, encoding: "base64" },
      });
      if (pptxResult._tag === "Failure")
        throw new Error("The deck source was saved, but the PowerPoint export failed.");
      pendingLocalSave.current = contents;
      lastExportedSource.current = contents;
      setProjectFileQueryData(environmentId, cwd, sourcePath, contents);
      confirmProjectFileQueryData(environmentId, cwd, sourcePath, contents);
      setSavedPath(sourcePath);
      setDirty(false);
      onSaved?.(sourcePath);
      if (download)
        downloadPresentation(base64, `${base.split("/").at(-1) ?? "Presentation"}.pptx`);
      return sourcePath;
    } finally {
      setBusy(false);
    }
  };
  const run = (download: boolean) => {
    void save(download).catch((cause) =>
      setError(cause instanceof Error ? cause.message : "Could not save presentation."),
    );
  };
  const askAi = async () => {
    const request = brief.trim();
    if (!request) return;
    try {
      const path = await save(false);
      const inserted = onAskAi?.(
        `${selectedElement ? `Slide ${selected + 1}, ${selectedElement} in ` : `For `}${path}: ${request}`,
      );
      if (!inserted) throw new Error("Open a project chat before asking AI.");
      setHistory((current) => [
        ...current,
        { role: "user", text: request },
        { role: "status", text: "Ready in project chat. Send it there." },
      ]);
      setBrief("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not prepare the deck for AI.");
    }
  };
  const moveSlide = (offset: number) => {
    const target = selected + offset;
    if (target < 0 || target >= deck.slides.length) return;
    setDeck((current) => {
      const slides = [...current.slides];
      [slides[selected], slides[target]] = [slides[target]!, slides[selected]!];
      return { ...current, slides };
    });
    setSelected(target);
  };

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden">
      <div className="flex flex-wrap items-center gap-2 border-b p-3">
        <Input
          aria-label="Presentation title"
          className="max-w-64"
          value={deck.title}
          onChange={(event) => editDeck({ title: event.target.value })}
        />
        <span className="flex-1" />
        <Button size="sm" variant="outline" disabled={busy} onClick={() => run(false)}>
          {busy ? "Saving..." : "Save to project"}
        </Button>
        <Button size="sm" disabled={busy} onClick={() => run(true)}>
          <DownloadIcon /> Save and download
        </Button>
      </div>
      {error ? (
        <p role="alert" className="bg-destructive/10 px-4 py-2 text-sm text-destructive">
          {error}
        </p>
      ) : null}
      <div className="grid min-h-0 flex-1 grid-cols-1 overflow-auto lg:grid-cols-[180px_minmax(0,1fr)_380px]">
        <nav
          aria-label="Slides"
          className="space-y-2 border-b p-3 lg:overflow-auto lg:border-b-0 lg:border-r"
        >
          {deck.slides.map((slide, index) => (
            <button
              key={slide.id}
              type="button"
              className={`w-full rounded-md border p-2 text-left text-xs ${selected === index ? "border-primary bg-accent" : "hover:bg-muted"}`}
              onClick={() => setSelected(index)}
            >
              <span className="text-muted-foreground">
                {index + 1} · {slide.layout}
              </span>
              <span className="mt-1 block truncate font-medium">{slide.title}</span>
            </button>
          ))}
          <Button
            size="sm"
            variant="outline"
            className="w-full"
            onClick={() => {
              setDeck((current) => ({ ...current, slides: [...current.slides, newSlide()] }));
              setSelected(deck.slides.length);
            }}
          >
            <PlusIcon /> Add slide
          </Button>
        </nav>
        <main className="min-w-0 space-y-4 p-5 lg:overflow-auto">
          {selectedSlide ? (
            <>
              <SlideCanvas
                slide={selectedSlide}
                index={selected}
                selecting={selecting}
                selectedElement={selectedElement}
                onSelectElement={(element) => {
                  setSelectedElement(element);
                  setSelecting(false);
                }}
              />
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="text-xs font-medium">
                  Layout
                  <select
                    className="mt-1 w-full rounded-md border bg-background p-2 text-sm"
                    value={selectedSlide.layout}
                    onChange={(event) => editSlide({ layout: event.target.value as SlideLayout })}
                  >
                    {LAYOUTS.map((layout) => (
                      <option value={layout.value} key={layout.value}>
                        {layout.label}
                      </option>
                    ))}
                  </select>
                </label>
                <div className="flex items-end gap-2">
                  <Button
                    size="sm"
                    variant="outline"
                    disabled={selected === 0}
                    onClick={() => moveSlide(-1)}
                    aria-label="Move slide up"
                  >
                    <ArrowUpIcon />
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    disabled={selected === deck.slides.length - 1}
                    onClick={() => moveSlide(1)}
                    aria-label="Move slide down"
                  >
                    <ArrowDownIcon />
                  </Button>
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => {
                      setDeck((current) => ({
                        ...current,
                        slides: current.slides.filter((_, index) => index !== selected),
                      }));
                      setSelected(Math.max(0, selected - 1));
                    }}
                    aria-label="Remove slide"
                  >
                    <Trash2Icon />
                  </Button>
                </div>
              </div>
              <Input
                aria-label="Slide title"
                value={selectedSlide.title}
                onChange={(event) => editSlide({ title: event.target.value })}
              />
              <textarea
                aria-label="Slide content"
                className="min-h-28 w-full rounded-md border bg-background p-3 text-sm"
                value={selectedSlide.body}
                onChange={(event) => editSlide({ body: event.target.value })}
                placeholder="Key points or chart insight"
              />
              {selectedSlide.layout === "two-column" ? (
                <textarea
                  aria-label="Right column"
                  className="min-h-28 w-full rounded-md border bg-background p-3 text-sm"
                  value={selectedSlide.rightBody ?? ""}
                  onChange={(event) => editSlide({ rightBody: event.target.value })}
                  placeholder="Right column"
                />
              ) : null}
              <Input
                aria-label="Source citation"
                value={selectedSlide.source ?? ""}
                onChange={(event) => editSlide({ source: event.target.value })}
                placeholder="Source file or citation"
              />
              {selectedSlide.layout === "chart" ? (
                <div className="space-y-2 rounded-md border p-3">
                  <label className="text-sm font-medium">
                    Chart data from Excel or CSV
                    <input
                      type="file"
                      accept=".xlsx,.csv,.tsv"
                      className="mt-2 block w-full text-xs"
                      onChange={(event) => {
                        const file = event.target.files?.[0];
                        if (file)
                          void importPresentationChart(file)
                            .then((chart) => {
                              editSlide({
                                chart,
                                ...(chart.source ? { source: chart.source } : {}),
                              });
                              setError("");
                            })
                            .catch((cause) =>
                              setError(
                                cause instanceof Error
                                  ? cause.message
                                  : "Could not read spreadsheet.",
                              ),
                            );
                      }}
                    />
                  </label>
                  <select
                    aria-label="Chart type"
                    className="rounded-md border bg-background p-2 text-sm"
                    value={selectedSlide.chart?.type ?? "bar"}
                    onChange={(event) => {
                      const type = event.target.value as PresentationChart["type"];
                      const current = selectedSlide.chart;
                      if (!current) return;
                      const baseSeries =
                        current.type === "scatter" && type !== "scatter"
                          ? current.series.slice(1)
                          : current.series;
                      if (
                        (type === "pie" || type === "doughnut") &&
                        baseSeries[0]?.values.some((value) => value < 0)
                      ) {
                        setError("Een cirkel- of ringdiagram vereist niet-negatieve waarden.");
                        return;
                      }
                      const series =
                        type === "scatter"
                          ? current.type === "scatter"
                            ? current.series
                            : baseSeries.length >= 2
                              ? baseSeries.slice(0, 2)
                              : [
                                  {
                                    name: "X",
                                    values: current.categories.map((_, index) => index + 1),
                                  },
                                  baseSeries[0]!,
                                ]
                          : type === "pie" || type === "doughnut"
                            ? baseSeries.slice(0, 1)
                            : baseSeries;
                      editSlide({ chart: { ...current, type, series } });
                      setError("");
                    }}
                  >
                    {CONSULTANCY_CHART_TYPES.map((component) => (
                      <option key={component.type} value={component.type}>
                        {component.name}
                      </option>
                    ))}
                  </select>
                  <Input
                    aria-label="Chart unit"
                    className="max-w-32"
                    placeholder="Eenheid, bijv. %"
                    value={selectedSlide.chart?.unit ?? ""}
                    onChange={(event) =>
                      editSlide({
                        chart: {
                          type: selectedSlide.chart?.type ?? "bar",
                          categories: selectedSlide.chart?.categories ?? [],
                          series: selectedSlide.chart?.series ?? [],
                          ...(selectedSlide.chart?.source
                            ? { source: selectedSlide.chart.source }
                            : {}),
                          unit: event.target.value,
                        },
                      })
                    }
                  />
                  <p className="text-xs text-muted-foreground">
                    First worksheet: labels in column A, numeric series in columns B through E. Edit
                    the source with AI for more complex data.
                  </p>
                </div>
              ) : null}
            </>
          ) : (
            <p className="text-sm text-muted-foreground">Add a slide to begin.</p>
          )}
        </main>
        <aside className="flex min-h-0 flex-col border-t bg-muted/20 lg:border-l lg:border-t-0">
          <div className="flex items-center justify-between border-b p-4">
            <div>
              <h3 className="text-sm font-semibold">Presentation chat</h3>
              <p className="text-xs text-muted-foreground">Ask a question or request a change</p>
            </div>
            <Button
              size="icon-sm"
              variant={selecting ? "default" : "outline"}
              aria-label="Select slide element"
              title="Select slide element"
              onClick={() => {
                setSelecting((value) => !value);
                setSelectedElement(null);
              }}
            >
              <MousePointer2Icon />
            </Button>
          </div>
          {projectId ? (
            <PresentationChat
              environmentId={environmentId}
              projectId={projectId}
              {...(deck.chatThreadId
                ? { threadRef: scopeThreadRef(environmentId, deck.chatThreadId) }
                : {})}
              editorContext={presentationAgentContext(
                savedPath ?? `${deckPath(deck.title)}.t3deck.json`,
                selected + 1,
                selectedElement,
              )}
              onThreadRefChange={setChatThreadRef}
              onBeforeStart={() => (savedPath && !dirty ? Promise.resolve(savedPath) : save(false))}
              onThreadCreated={async (path, threadId) => {
                const updated = { ...deck, chatThreadId: threadId };
                const contents = JSON.stringify(updated, null, 2);
                const result = await writeFile({
                  environmentId,
                  input: { cwd, relativePath: path, contents },
                });
                if (result._tag === "Failure")
                  throw new Error("Could not link the chat to this presentation.");
                setDeck(updated);
                pendingLocalSave.current = contents;
                setProjectFileQueryData(environmentId, cwd, path, contents);
                confirmProjectFileQueryData(environmentId, cwd, path, contents);
              }}
            />
          ) : (
            <div
              className="min-h-32 flex-1 space-y-3 overflow-auto p-4"
              aria-label="Presentation requests"
            >
              {history.length === 0 ? (
                <p className="text-sm text-muted-foreground">Your requests will appear here.</p>
              ) : (
                history.map((message, index) => (
                  <p
                    key={index}
                    className={`rounded-lg p-3 text-sm ${message.role === "user" ? "bg-primary/10" : "bg-muted text-muted-foreground"}`}
                  >
                    {message.text}
                  </p>
                ))
              )}
            </div>
          )}
          {!projectId ? (
            <section className="space-y-2 border-t p-4">
              {selecting ? (
                <p className="text-xs text-muted-foreground">
                  Click a title, text block, or chart on the slide.
                </p>
              ) : null}
              {selectedElement ? (
                <button
                  type="button"
                  className="rounded-full border px-2 py-1 text-xs"
                  onClick={() => setSelectedElement(null)}
                >
                  Slide {selected + 1} · {selectedElement} ×
                </button>
              ) : null}
              <textarea
                aria-label="AI presentation request"
                className="min-h-28 w-full rounded-md border bg-background p-2 text-sm"
                value={brief}
                onChange={(event) => setBrief(event.target.value)}
                placeholder="Describe a change or ask a question…"
              />
              <Button
                size="sm"
                variant="outline"
                className="w-full"
                disabled={busy || !onAskAi || !brief.trim()}
                onClick={() => void askAi()}
              >
                <SparklesIcon /> Add to project chat
              </Button>
            </section>
          ) : null}
        </aside>
      </div>
    </div>
  );
}

export function PresentationMaker(props: {
  environmentId: EnvironmentId;
  cwd: string;
  onOpenInChat?: (prompt: string) => boolean;
  onSaved?: (path: string) => void;
  buttonLabel?: string;
  projectId?: ProjectId;
}) {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(newDeck);
  return (
    <>
      <Button
        type="button"
        size={props.buttonLabel ? "sm" : "icon-xs"}
        variant="ghost"
        aria-label="Create presentation"
        onClick={() => {
          setDraft(newDeck());
          setOpen(true);
        }}
      >
        <PresentationIcon className="size-3.5" />
        {props.buttonLabel}
      </Button>
      <Dialog open={open} onOpenChange={setOpen}>
        <DialogPopup className="flex h-[min(92vh,850px)] w-[min(1300px,96vw)] max-w-none flex-col overflow-hidden">
          <DialogTitle>Presentation studio</DialogTitle>
          <PresentationEditor
            key={draft.slides[0]?.id}
            environmentId={props.environmentId}
            cwd={props.cwd}
            initialDeck={draft}
            {...(props.projectId ? { projectId: props.projectId } : {})}
            onSaved={(path) => {
              props.onSaved?.(path);
            }}
            onAskAi={(prompt) => {
              const inserted = props.onOpenInChat?.(prompt) ?? false;
              return inserted;
            }}
          />
        </DialogPopup>
      </Dialog>
    </>
  );
}

export function PresentationDeckPreview(props: {
  contents: string;
  name: string;
  environmentId: EnvironmentId;
  cwd: string;
  onAskAi?: (prompt: string) => boolean;
  projectId?: ProjectId;
}) {
  const parsed = useMemo(() => {
    try {
      return { deck: parseDeck(props.contents), error: null };
    } catch (cause) {
      return {
        deck: null,
        error: cause instanceof Error ? cause.message : "Could not open this presentation.",
      };
    }
  }, [props.contents]);
  if (!parsed.deck)
    return (
      <div role="alert" className="p-6 text-sm text-destructive">
        {parsed.error}
      </div>
    );
  return (
    <PresentationEditor
      key={props.name}
      environmentId={props.environmentId}
      cwd={props.cwd}
      relativePath={props.name}
      initialDeck={parsed.deck}
      {...(props.projectId ? { projectId: props.projectId } : {})}
      {...(props.onAskAi ? { onAskAi: props.onAskAi } : {})}
    />
  );
}
