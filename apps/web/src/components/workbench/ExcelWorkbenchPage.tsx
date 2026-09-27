import { Badge } from "~/components/ui/badge";
import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import {
  Select,
  SelectItem,
  SelectPopup,
  SelectTrigger,
  SelectValue,
} from "~/components/ui/select";
import { WorkspacePageHeader } from "~/components/WorkspacePageHeader";
import {
  AlertCircleIcon,
  CheckCircle2Icon,
  ChevronLeftIcon,
  ChevronRightIcon,
  DatabaseIcon,
  DownloadIcon,
  FileSpreadsheetIcon,
  ShieldCheckIcon,
  UploadIcon,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";

import {
  formatWorkbookCell,
  formatWorkbookNumber,
  WORKBOOK_PAGE_SIZE,
  type WorkbookCell,
  type WorkbookChartPoint,
  type WorkbookCommand,
  type WorkbookColumnProfile,
  type WorkbookResponse,
  type WorkbookSheetOverview,
} from "./excelWorkbenchModel";

interface PendingRequest {
  readonly resolve: (response: Exclude<WorkbookResponse, { type: "error" }>) => void;
  readonly reject: (error: Error) => void;
}

export function ExcelWorkbenchPage({
  initialFile,
  onBack,
}: {
  initialFile?: File | null;
  onBack?: () => void;
}) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const workerRef = useRef<Worker | null>(null);
  const pendingRequestsRef = useRef(new Map<number, PendingRequest>());
  const requestIdRef = useRef(0);
  const latestQueryRef = useRef(0);
  const latestAnswerRef = useRef(0);
  const latestChartRef = useRef(0);
  const [fileName, setFileName] = useState("");
  const [fileSize, setFileSize] = useState(0);
  const [sheetNames, setSheetNames] = useState<ReadonlyArray<string>>([]);
  const [overview, setOverview] = useState<WorkbookSheetOverview | null>(null);
  const [rows, setRows] = useState<ReadonlyArray<ReadonlyArray<WorkbookCell>>>([]);
  const [filteredRowCount, setFilteredRowCount] = useState(0);
  const [search, setSearch] = useState("");
  const [page, setPage] = useState(0);
  const [selectedColumn, setSelectedColumn] = useState("");
  const [groupColumn, setGroupColumn] = useState("");
  const [valueColumn, setValueColumn] = useState("");
  const [chartPoints, setChartPoints] = useState<ReadonlyArray<WorkbookChartPoint>>([]);
  const [chartMessage, setChartMessage] = useState("");
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [isAsking, setIsAsking] = useState(false);
  const [isCharting, setIsCharting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [exportMessage, setExportMessage] = useState("");

  const getWorker = useCallback(() => {
    if (workerRef.current) return workerRef.current;
    const worker = new Worker(new URL("./excelWorkbench.worker.ts", import.meta.url), {
      type: "module",
    });
    worker.onmessage = (event: MessageEvent<WorkbookResponse>) => {
      const response = event.data;
      const pending = pendingRequestsRef.current.get(response.requestId);
      if (!pending) return;
      pendingRequestsRef.current.delete(response.requestId);
      if (response.type === "error") pending.reject(new Error(response.message));
      else pending.resolve(response);
    };
    worker.onerror = (event) => {
      const error = new Error(event.message || "The workbook processor stopped unexpectedly.");
      for (const pending of pendingRequestsRef.current.values()) pending.reject(error);
      pendingRequestsRef.current.clear();
      worker.terminate();
      workerRef.current = null;
    };
    workerRef.current = worker;
    return worker;
  }, []);

  const sendCommand = useCallback(
    (command: WorkbookCommand) =>
      new Promise<Exclude<WorkbookResponse, { type: "error" }>>((resolve, reject) => {
        const requestId = ++requestIdRef.current;
        pendingRequestsRef.current.set(requestId, { resolve, reject });
        getWorker().postMessage({ ...command, requestId });
      }),
    [getWorker],
  );

  useEffect(
    () => () => {
      workerRef.current?.terminate();
      for (const pending of pendingRequestsRef.current.values()) {
        pending.reject(new Error("The workbench was closed."));
      }
      pendingRequestsRef.current.clear();
    },
    [],
  );

  const applyLoadedResponse = useCallback(
    (response: Extract<WorkbookResponse, { type: "loaded" | "sheet" }>) => {
      latestAnswerRef.current++;
      latestChartRef.current++;
      setFileName(response.fileName);
      setFileSize(response.fileSize);
      setSheetNames(response.sheetNames);
      setOverview(response.overview);
      setRows(response.rows);
      setFilteredRowCount(response.filteredRowCount);
      setPage(response.page);
      setSelectedColumn(response.overview.columns[0]?.name ?? "");
      setGroupColumn(
        response.overview.columns.find((column) => column.type === "Text")?.name ??
          response.overview.columns.find((column) => column.type === "Date")?.name ??
          "",
      );
      setValueColumn(
        response.overview.columns.find((column) => column.type === "Number")?.name ?? "",
      );
      setChartPoints([]);
      setChartMessage("");
      setQuestion("");
      setAnswer("");
      setIsAsking(false);
      setIsCharting(false);
      setSearch("");
      setErrorMessage("");
      setExportMessage("");
    },
    [],
  );

  const loadFile = useCallback(
    async (file: File | undefined) => {
      if (!file) return;
      setIsImporting(true);
      setErrorMessage("");
      setExportMessage("");
      try {
        const response = await sendCommand({ type: "load", file });
        if (response.type === "loaded") applyLoadedResponse(response);
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "The workbook could not be opened.",
        );
      } finally {
        setIsImporting(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    },
    [applyLoadedResponse, sendCommand],
  );

  useEffect(() => {
    if (initialFile) void loadFile(initialFile);
  }, [initialFile, loadFile]);

  const selectSheet = useCallback(
    async (name: string) => {
      setIsImporting(true);
      setErrorMessage("");
      try {
        const response = await sendCommand({ type: "sheet", name });
        if (response.type === "sheet") applyLoadedResponse(response);
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "The worksheet could not be opened.",
        );
      } finally {
        setIsImporting(false);
      }
    },
    [applyLoadedResponse, sendCommand],
  );

  useEffect(() => {
    if (!overview) return;
    const queryId = ++latestQueryRef.current;
    const timeout = window.setTimeout(
      () => {
        void sendCommand({ type: "query", term: search, page })
          .then((response) => {
            if (queryId !== latestQueryRef.current || response.type !== "rows") return;
            setRows(response.rows);
            setFilteredRowCount(response.filteredRowCount);
          })
          .catch((error: unknown) => {
            if (queryId === latestQueryRef.current) {
              setErrorMessage(
                error instanceof Error ? error.message : "Could not search this worksheet.",
              );
            }
          });
      },
      search ? 180 : 0,
    );
    return () => window.clearTimeout(timeout);
  }, [overview, page, search, sendCommand]);

  const askQuestion = useCallback(
    async (value: string) => {
      const normalizedQuestion = value.trim();
      if (!normalizedQuestion) return;
      const answerId = ++latestAnswerRef.current;
      setIsAsking(true);
      setAnswer("");
      try {
        const response = await sendCommand({ type: "ask", question: normalizedQuestion });
        if (answerId === latestAnswerRef.current && response.type === "answer") {
          setAnswer(response.text);
        }
      } catch (error) {
        if (answerId === latestAnswerRef.current) {
          setAnswer(error instanceof Error ? error.message : "Could not answer that question.");
        }
      } finally {
        if (answerId === latestAnswerRef.current) setIsAsking(false);
      }
    },
    [sendCommand],
  );

  useEffect(() => {
    if (!overview || !groupColumn || !valueColumn) {
      setChartPoints([]);
      setIsCharting(false);
      return;
    }
    const chartId = ++latestChartRef.current;
    setChartMessage("");
    setIsCharting(true);
    void sendCommand({ type: "chart", groupColumn, valueColumn })
      .then((response) => {
        if (chartId !== latestChartRef.current || response.type !== "chart") return;
        setChartPoints(response.points);
      })
      .catch((error: unknown) => {
        if (chartId !== latestChartRef.current) return;
        setChartMessage(
          error instanceof Error ? error.message : "Could not summarize these fields.",
        );
      })
      .finally(() => {
        if (chartId === latestChartRef.current) setIsCharting(false);
      });
  }, [groupColumn, overview, sendCommand, valueColumn]);

  const exportWorkbook = useCallback(async () => {
    setIsExporting(true);
    setErrorMessage("");
    setExportMessage("");
    try {
      const response = await sendCommand({ type: "export" });
      if (response.type !== "exported") return;
      const url = URL.createObjectURL(
        new Blob([response.data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      );
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = response.fileName;
      anchor.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
      setExportMessage(`${response.fileName} downloaded`);
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "The analysis workbook could not be exported.",
      );
    } finally {
      setIsExporting(false);
    }
  }, [sendCommand]);

  const selectedProfile =
    overview?.columns.find((column) => column.name === selectedColumn) ?? null;
  const numericColumns = overview?.columns.filter((column) => column.type === "Number") ?? [];
  const categoryColumns = overview?.columns.filter((column) => column.type !== "Number") ?? [];
  const totalPages = Math.max(1, Math.ceil(filteredRowCount / WORKBOOK_PAGE_SIZE));

  return (
    <main className="flex h-dvh min-h-0 flex-col overflow-hidden overscroll-y-none bg-background">
      <WorkspacePageHeader className="border-b border-border/70">
        {onBack ? (
          <Button size="sm" variant="ghost" onClick={onBack}>
            <ChevronLeftIcon /> Project
          </Button>
        ) : null}
        <div className="flex min-w-0 items-center gap-3">
          <div className="flex size-8 shrink-0 items-center justify-center rounded-lg border border-border/70 bg-card text-primary">
            <FileSpreadsheetIcon className="size-4" />
          </div>
          <div className="min-w-0">
            <p className="truncate text-sm font-medium text-foreground">Customer data</p>
            <p className="text-xs text-muted-foreground">Excel workbench</p>
          </div>
        </div>
        <div className="ml-auto flex items-center gap-2">
          <span className="hidden sm:inline-flex">
            <Badge variant="success" size="sm">
              <ShieldCheckIcon />
              Processed in browser
            </Badge>
          </span>
          {overview ? (
            <>
              <Button
                disabled={isExporting || isImporting}
                onClick={() => void exportWorkbook()}
                size="sm"
                variant="outline"
              >
                <DownloadIcon />
                {isExporting ? "Preparing…" : "Export analysis"}
              </Button>
              <Button
                disabled={isImporting}
                onClick={() => fileInputRef.current?.click()}
                size="sm"
              >
                <UploadIcon />
                Replace workbook
              </Button>
            </>
          ) : null}
        </div>
      </WorkspacePageHeader>

      <input
        ref={fileInputRef}
        accept=".xlsx,.csv,.tsv"
        className="sr-only"
        disabled={isImporting}
        onChange={(event) => void loadFile(event.currentTarget.files?.[0])}
        type="file"
      />

      {!overview ? (
        <main className="min-h-0 flex-1 overflow-y-auto">
          <div className="mx-auto flex min-h-full w-full max-w-6xl flex-col justify-center px-5 py-10 sm:px-8">
            <div className="mb-8 max-w-2xl">
              <div className="mb-4">
                <Badge variant="outline">
                  <DatabaseIcon />
                  Client data analysis
                </Badge>
              </div>
              <h1 className="text-3xl font-semibold tracking-tight text-foreground sm:text-4xl">
                Make sense of your customer data.
              </h1>
              <p className="mt-3 max-w-xl text-sm leading-6 text-muted-foreground sm:text-base">
                Open an Excel or delimited file to review data quality, compare customer segments,
                ask quick questions, and export an analysis workbook.
              </p>
            </div>

            <div
              aria-disabled={isImporting}
              aria-label="Choose a customer workbook"
              className={`group flex min-h-64 cursor-pointer flex-col items-center justify-center rounded-2xl border border-dashed px-6 py-10 text-center transition-colors ${dragging ? "border-primary bg-primary/6" : "border-border/80 bg-card/30 hover:border-foreground/25 hover:bg-card/55"}`}
              onClick={() => {
                if (!isImporting) fileInputRef.current?.click();
              }}
              onDragEnter={(event) => {
                event.preventDefault();
                setDragging(true);
              }}
              onDragLeave={(event) => {
                if (
                  !(event.relatedTarget instanceof Node) ||
                  !event.currentTarget.contains(event.relatedTarget)
                ) {
                  setDragging(false);
                }
              }}
              onDragOver={(event) => event.preventDefault()}
              onKeyDown={(event) => {
                if ((event.key === "Enter" || event.key === " ") && !isImporting) {
                  event.preventDefault();
                  fileInputRef.current?.click();
                }
              }}
              onDrop={(event) => {
                event.preventDefault();
                setDragging(false);
                if (!isImporting) void loadFile(event.dataTransfer.files[0]);
              }}
              role="button"
              tabIndex={0}
            >
              <span className="mb-4 flex size-12 items-center justify-center rounded-xl border border-border/70 bg-background text-primary shadow-sm/5 group-hover:bg-accent/50">
                {isImporting ? (
                  <span className="size-5 animate-spin rounded-full border-2 border-primary/25 border-t-primary" />
                ) : (
                  <UploadIcon className="size-5" />
                )}
              </span>
              <span className="text-base font-medium text-foreground">
                {isImporting ? "Reading workbook…" : "Drop a customer workbook here"}
              </span>
              <span className="mt-1 text-sm text-muted-foreground">
                or <span className="text-primary underline underline-offset-4">browse files</span>
              </span>
              <span className="mt-4 text-xs text-muted-foreground">
                XLSX, CSV, or TSV · up to 20 MB
              </span>
            </div>

            {errorMessage ? <ErrorNotice message={errorMessage} /> : null}

            <div className="mt-8 grid gap-3 sm:grid-cols-3">
              <FeatureHint
                title="Check data quality"
                description="Find empty fields, duplicate rows, and inconsistent columns."
              />
              <FeatureHint
                title="Compare customer segments"
                description="Summarize a numeric field by region, status, or another category."
              />
              <FeatureHint
                title="Get quick answers"
                description="Ask about every sheet, column types, missing values, duplicates, or numeric totals."
              />
            </div>
            <p className="mt-6 flex items-center justify-center gap-2 text-xs text-muted-foreground">
              <ShieldCheckIcon className="size-3.5" />
              Your workbook stays in this browser. It is not uploaded to the T3 server.
            </p>
          </div>
        </main>
      ) : (
        <main className="min-h-0 flex-1 overflow-y-auto">
          <div className="mx-auto flex w-full max-w-[1480px] flex-col gap-5 px-5 py-6 sm:px-8">
            <section className="flex flex-wrap items-start justify-between gap-4">
              <div className="min-w-0">
                <p className="text-xs font-medium uppercase tracking-wider text-muted-foreground">
                  Workbook
                </p>
                <h1 className="mt-1 truncate text-xl font-semibold tracking-tight text-foreground sm:text-2xl">
                  {fileName}
                </h1>
                <p className="mt-1 text-sm text-muted-foreground">
                  {formatFileSize(fileSize)} <span className="px-1.5 text-border">·</span>{" "}
                  {overview.name}
                </p>
              </div>
              <span className="sm:hidden">
                <Badge variant="success" size="sm">
                  <ShieldCheckIcon />
                  Local analysis
                </Badge>
              </span>
            </section>

            {errorMessage ? <ErrorNotice message={errorMessage} /> : null}
            {exportMessage ? (
              <p
                aria-live="polite"
                className="flex items-center gap-2 text-sm text-success-foreground"
              >
                <CheckCircle2Icon className="size-4" />
                {exportMessage}
              </p>
            ) : null}

            <section aria-label="Workbook summary" className="grid gap-3 sm:grid-cols-3">
              <SummaryCard
                label="Customer records"
                value={overview.rowCount}
                detail="Rows with data"
                icon={<DatabaseIcon />}
              />
              <SummaryCard
                label="Fields"
                value={overview.columns.length}
                detail="Columns in this sheet"
                icon={<FileSpreadsheetIcon />}
              />
              <SummaryCard
                label="Data completeness"
                value={`${String(overview.completeness)}%`}
                detail={`${formatWorkbookNumber(overview.missingCells)} empty cells`}
                icon={overview.completeness >= 95 ? <CheckCircle2Icon /> : <AlertCircleIcon />}
                tone={overview.completeness >= 95 ? "good" : "warning"}
              />
            </section>

            {sheetNames.length > 1 ? (
              <nav
                aria-label="Workbook worksheets"
                className="flex min-w-0 gap-1 overflow-x-auto border-b border-border/70"
              >
                {sheetNames.map((name) => (
                  <button
                    aria-current={name === overview.name ? "page" : undefined}
                    className={`relative -mb-px shrink-0 px-3 py-2.5 text-sm transition-colors ${name === overview.name ? "border-b-2 border-primary font-medium text-foreground" : "text-muted-foreground hover:text-foreground"}`}
                    disabled={isImporting}
                    key={name}
                    onClick={() => void selectSheet(name)}
                    type="button"
                  >
                    {name}
                  </button>
                ))}
              </nav>
            ) : null}

            <section className="grid min-w-0 gap-4 xl:grid-cols-[minmax(0,1.8fr)_minmax(300px,0.8fr)]">
              <div className="min-w-0 overflow-hidden rounded-xl border border-border/70 bg-card/30">
                <div className="flex flex-wrap items-center justify-between gap-3 border-b border-border/70 px-4 py-3">
                  <div>
                    <h2 className="text-sm font-semibold text-foreground">Data preview</h2>
                    <p aria-live="polite" className="mt-0.5 text-xs text-muted-foreground">
                      {formatWorkbookNumber(filteredRowCount)} {search ? "matching" : "total"}{" "}
                      records
                    </p>
                  </div>
                  <div className="w-full sm:w-64">
                    <Input
                      aria-label="Search worksheet rows"
                      disabled={isImporting}
                      onChange={(event) => {
                        setSearch(event.currentTarget.value);
                        setPage(0);
                      }}
                      placeholder="Search rows…"
                      value={search}
                    />
                  </div>
                </div>
                {overview.columns.length > 0 ? (
                  <div className="max-h-[min(58vh,620px)] min-h-72 overflow-auto">
                    <table className="w-full min-w-max border-collapse text-xs">
                      <thead className="sticky top-0 z-10 bg-muted/95 backdrop-blur-sm">
                        <tr className="border-b border-border/70">
                          <th className="sticky left-0 z-20 w-12 bg-muted px-3 py-2.5 text-right font-medium text-muted-foreground">
                            #
                          </th>
                          {overview.columns.map((column) => (
                            <th
                              className="max-w-64 px-3 py-2.5 text-left font-medium text-muted-foreground"
                              key={column.name}
                            >
                              <span className="block truncate">{column.name}</span>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((row, rowIndex) => (
                          <tr
                            className="border-b border-border/40 last:border-0 hover:bg-muted/30"
                            key={`${page}-${rowIndex}`}
                          >
                            <td className="sticky left-0 bg-background/95 px-3 py-2 text-right tabular-nums text-muted-foreground">
                              {formatWorkbookNumber(page * WORKBOOK_PAGE_SIZE + rowIndex + 1)}
                            </td>
                            {overview.columns.map((column, columnIndex) => {
                              const value = row[columnIndex] ?? null;
                              return (
                                <td
                                  className="max-w-64 px-3 py-2 text-foreground"
                                  key={column.name}
                                >
                                  <span className="block truncate">
                                    {formatWorkbookCell(value) || (
                                      <span className="text-muted-foreground/45">—</span>
                                    )}
                                  </span>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                        {rows.length === 0 ? (
                          <tr>
                            <td
                              className="px-4 py-12 text-center text-sm text-muted-foreground"
                              colSpan={overview.columns.length + 1}
                            >
                              {search
                                ? "No rows match that search."
                                : "This worksheet has no data rows."}
                            </td>
                          </tr>
                        ) : null}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="flex min-h-72 items-center justify-center px-6 text-center text-sm text-muted-foreground">
                    This worksheet is empty or does not have a header row.
                  </div>
                )}
                <div className="flex items-center justify-between border-t border-border/70 px-4 py-2.5">
                  <p className="text-xs text-muted-foreground">
                    {filteredRowCount === 0
                      ? "0 records"
                      : `Page ${String(page + 1)} of ${String(totalPages)}`}
                  </p>
                  <div className="flex items-center gap-1">
                    <Button
                      aria-label="Previous page"
                      disabled={page === 0 || isImporting}
                      onClick={() => setPage((current) => Math.max(0, current - 1))}
                      size="icon-sm"
                      variant="ghost"
                    >
                      <ChevronLeftIcon />
                    </Button>
                    <Button
                      aria-label="Next page"
                      disabled={page + 1 >= totalPages || isImporting}
                      onClick={() => setPage((current) => Math.min(totalPages - 1, current + 1))}
                      size="icon-sm"
                      variant="ghost"
                    >
                      <ChevronRightIcon />
                    </Button>
                  </div>
                </div>
              </div>

              <aside className="flex min-w-0 flex-col gap-4">
                <section className="rounded-xl border border-border/70 bg-card/30 p-4">
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <h2 className="text-sm font-semibold text-foreground">Data quality</h2>
                      <p className="mt-1 text-xs leading-5 text-muted-foreground">
                        {overview.missingCells === 0 && overview.duplicateRows === 0
                          ? "No empty cells or duplicate records found."
                          : "Review these patterns before using the data."}
                      </p>
                    </div>
                    <Badge
                      variant={
                        overview.missingCells === 0 && overview.duplicateRows === 0
                          ? "success"
                          : "warning"
                      }
                    >
                      {overview.completeness}% complete
                    </Badge>
                  </div>
                  <div className="mt-4 space-y-3">
                    <QualityMetric
                      label="Empty cells"
                      value={overview.missingCells}
                      total={overview.rowCount * overview.columns.length}
                    />
                    <QualityMetric
                      label="Repeated records"
                      value={overview.duplicateRows}
                      total={overview.rowCount}
                    />
                  </div>
                </section>

                <section className="min-w-0 rounded-xl border border-border/70 bg-card/30 p-4">
                  <div className="mb-3">
                    <h2 className="text-sm font-semibold text-foreground">Quick answers</h2>
                    <p className="mt-1 text-xs text-muted-foreground">
                      Ask about every sheet, column types, data gaps, duplicates, or numeric totals.
                    </p>
                  </div>
                  <form
                    className="flex items-center gap-2"
                    onSubmit={(event) => {
                      event.preventDefault();
                      void askQuestion(question);
                    }}
                  >
                    <Input
                      aria-label="Ask a question about this worksheet"
                      disabled={isImporting || isAsking}
                      onChange={(event) => setQuestion(event.currentTarget.value)}
                      placeholder="e.g. Average order value?"
                      value={question}
                    />
                    <Button disabled={isImporting || isAsking || !question.trim()} size="sm">
                      {isAsking ? "…" : "Ask"}
                    </Button>
                  </form>
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {["How many customer records?", "Which fields have missing values?"].map(
                      (suggestion) => (
                        <Button
                          disabled={isImporting || isAsking}
                          key={suggestion}
                          onClick={() => {
                            setQuestion(suggestion);
                            void askQuestion(suggestion);
                          }}
                          size="xs"
                          type="button"
                          variant="ghost"
                        >
                          {suggestion}
                        </Button>
                      ),
                    )}
                  </div>
                  {answer ? (
                    <p
                      aria-live="polite"
                      className="mt-3 rounded-lg bg-background/70 px-3 py-2.5 text-xs leading-5 text-foreground"
                    >
                      {answer}
                    </p>
                  ) : null}
                </section>

                <section className="min-w-0 rounded-xl border border-border/70 bg-card/30 p-4">
                  <div className="mb-3">
                    <h2 className="text-sm font-semibold text-foreground">Segment summary</h2>
                    <p className="mt-1 text-xs text-muted-foreground">
                      Add up a numeric field for each category.
                    </p>
                  </div>
                  {numericColumns.length > 0 ? (
                    <div className="grid gap-2">
                      <label className="flex min-w-0 flex-col gap-1 text-2xs text-muted-foreground">
                        Group by
                        <Select
                          value={groupColumn || null}
                          onValueChange={(value: string | null) => setGroupColumn(value ?? "")}
                        >
                          <SelectTrigger aria-label="Group totals by category">
                            <SelectValue placeholder="Choose a field" />
                          </SelectTrigger>
                          <SelectPopup>
                            {categoryColumns.map((column) => (
                              <SelectItem key={column.name} value={column.name}>
                                {column.name}
                              </SelectItem>
                            ))}
                          </SelectPopup>
                        </Select>
                      </label>
                      <label className="flex min-w-0 flex-col gap-1 text-2xs text-muted-foreground">
                        Sum
                        <Select
                          value={valueColumn || null}
                          onValueChange={(value: string | null) => setValueColumn(value ?? "")}
                        >
                          <SelectTrigger aria-label="Numeric field to sum">
                            <SelectValue placeholder="Choose a field" />
                          </SelectTrigger>
                          <SelectPopup>
                            {numericColumns.map((column) => (
                              <SelectItem key={column.name} value={column.name}>
                                {column.name}
                              </SelectItem>
                            ))}
                          </SelectPopup>
                        </Select>
                      </label>
                    </div>
                  ) : (
                    <p className="rounded-lg bg-background/70 px-3 py-2.5 text-xs leading-5 text-muted-foreground">
                      This sheet has no numeric fields to summarize.
                    </p>
                  )}
                  {chartMessage ? (
                    <p className="mt-3 text-xs text-destructive-foreground">{chartMessage}</p>
                  ) : null}
                  {!chartMessage && numericColumns.length > 0 ? (
                    <div aria-live="polite" className="mt-3 space-y-2">
                      {isCharting ? (
                        <p className="text-xs text-muted-foreground">Updating summary…</p>
                      ) : chartPoints.length > 0 ? (
                        <CategoryTotals points={chartPoints} />
                      ) : (
                        <p className="text-xs leading-5 text-muted-foreground">
                          {groupColumn
                            ? "No category values are available for these fields."
                            : "Add a text or date field to group these totals."}
                        </p>
                      )}
                    </div>
                  ) : null}
                </section>

                <section className="min-w-0 rounded-xl border border-border/70 bg-card/30 p-4">
                  <div className="mb-3">
                    <h2 className="text-sm font-semibold text-foreground">Column profile</h2>
                    <p className="mt-1 text-xs text-muted-foreground">
                      Select a field to inspect its values.
                    </p>
                  </div>
                  <div className="max-h-48 space-y-1 overflow-y-auto pr-1">
                    {overview.columns.map((column) => (
                      <button
                        aria-pressed={column.name === selectedColumn}
                        className={`flex w-full items-center justify-between gap-3 rounded-md px-2.5 py-2 text-left text-xs transition-colors ${column.name === selectedColumn ? "bg-accent text-foreground" : "text-muted-foreground hover:bg-muted/60 hover:text-foreground"}`}
                        key={column.name}
                        onClick={() => setSelectedColumn(column.name)}
                        type="button"
                      >
                        <span className="min-w-0 truncate font-medium">{column.name}</span>
                        <span className="shrink-0">
                          {column.missing === 0 ? "Complete" : `${String(column.missing)} empty`}
                        </span>
                      </button>
                    ))}
                  </div>
                  {selectedProfile ? (
                    <SelectedColumnProfile profile={selectedProfile} rowCount={overview.rowCount} />
                  ) : null}
                </section>
              </aside>
            </section>
            <p className="flex items-center justify-center gap-2 pb-2 text-xs text-muted-foreground">
              <ShieldCheckIcon className="size-3.5" />
              Analysis runs locally in your browser. The source workbook is never changed.
            </p>
          </div>
        </main>
      )}
    </main>
  );
}

function CategoryTotals({ points }: { readonly points: ReadonlyArray<WorkbookChartPoint> }) {
  const largestTotal = Math.max(...points.map((point) => Math.abs(point.total)), 0);
  return (
    <ol aria-label="Largest category totals" className="space-y-2.5">
      {points.map((point) => {
        const width = largestTotal === 0 ? 0 : (Math.abs(point.total) / largestTotal) * 100;
        return (
          <li key={point.label}>
            <div className="mb-1 flex items-center justify-between gap-3 text-xs">
              <span className="min-w-0 truncate text-foreground" aria-label={point.label}>
                {point.label}
              </span>
              <span className="shrink-0 font-medium tabular-nums text-foreground">
                {formatWorkbookNumber(point.total)}
              </span>
            </div>
            <div aria-hidden="true" className="h-1.5 overflow-hidden rounded-full bg-muted">
              <div
                className="h-full rounded-full bg-primary/70"
                style={{ width: `${String(width)}%` }}
              />
            </div>
            <p className="mt-0.5 text-2xs text-muted-foreground">
              {formatWorkbookNumber(point.records)} records
            </p>
          </li>
        );
      })}
    </ol>
  );
}

function FeatureHint({
  title,
  description,
}: {
  readonly title: string;
  readonly description: string;
}) {
  return (
    <div className="rounded-xl border border-border/60 bg-card/20 px-4 py-3.5">
      <h2 className="text-sm font-medium text-foreground">{title}</h2>
      <p className="mt-1 text-xs leading-5 text-muted-foreground">{description}</p>
    </div>
  );
}

function SummaryCard({
  label,
  value,
  detail,
  icon,
  tone = "default",
}: {
  readonly label: string;
  readonly value: number | string;
  readonly detail: string;
  readonly icon: ReactNode;
  readonly tone?: "default" | "good" | "warning";
}) {
  const iconClass =
    tone === "good"
      ? "text-success-foreground"
      : tone === "warning"
        ? "text-warning-foreground"
        : "text-primary";
  return (
    <div className="flex min-w-0 items-start justify-between gap-3 rounded-xl border border-border/70 bg-card/30 p-4">
      <div className="min-w-0">
        <p className="text-xs text-muted-foreground">{label}</p>
        <p className="mt-1 text-2xl font-semibold tabular-nums tracking-tight text-foreground">
          {formatMetric(value)}
        </p>
        <p className="mt-1 truncate text-xs text-muted-foreground">{detail}</p>
      </div>
      <span
        className={`flex size-8 shrink-0 items-center justify-center rounded-lg bg-background text-sm ${iconClass}`}
      >
        {icon}
      </span>
    </div>
  );
}

function QualityMetric({
  label,
  value,
  total,
}: {
  readonly label: string;
  readonly value: number;
  readonly total: number;
}) {
  const ratio = total > 0 ? Math.min(100, (value / total) * 100) : 0;
  return (
    <div>
      <div className="mb-1.5 flex items-center justify-between gap-3 text-xs">
        <span className="text-muted-foreground">{label}</span>
        <span className="font-medium tabular-nums text-foreground">
          {formatWorkbookNumber(value)}
        </span>
      </div>
      <div aria-hidden="true" className="h-1.5 overflow-hidden rounded-full bg-muted">
        <div className="h-full rounded-full bg-primary/70" style={{ width: `${String(ratio)}%` }} />
      </div>
    </div>
  );
}

function SelectedColumnProfile({
  profile,
  rowCount,
}: {
  readonly profile: WorkbookColumnProfile;
  readonly rowCount: number;
}) {
  return (
    <div className="mt-4 border-t border-border/60 pt-3">
      <div className="flex items-center justify-between gap-3">
        <div className="min-w-0">
          <p className="truncate text-sm font-medium text-foreground">{profile.name}</p>
          <p className="mt-0.5 text-xs text-muted-foreground">
            {profile.type} · {formatWorkbookNumber(profile.unique)} unique values
          </p>
        </div>
        <Badge variant="outline">
          {rowCount === 0 ? 0 : Math.round((profile.nonEmpty / rowCount) * 100)}% filled
        </Badge>
      </div>
      {profile.type === "Number" && profile.average !== null ? (
        <dl className="mt-3 grid grid-cols-2 gap-2">
          <ProfileStat label="Total" value={profile.total} />
          <ProfileStat label="Average" value={profile.average} />
          <ProfileStat label="Minimum" value={profile.minimum} />
          <ProfileStat label="Maximum" value={profile.maximum} />
        </dl>
      ) : (
        <div className="mt-3 space-y-2 text-xs">
          <div className="flex items-center justify-between gap-3 text-muted-foreground">
            <span>Non-empty records</span>
            <span className="tabular-nums text-foreground">
              {formatWorkbookNumber(profile.nonEmpty)}
            </span>
          </div>
          {profile.examples.length > 0 ? (
            <div className="flex flex-wrap gap-1.5">
              {profile.examples.map((example) => (
                <Badge key={example} variant="secondary" size="sm">
                  <span className="max-w-56 truncate">{example}</span>
                </Badge>
              ))}
            </div>
          ) : (
            <p className="text-muted-foreground">No values in this field.</p>
          )}
        </div>
      )}
    </div>
  );
}

function ProfileStat({ label, value }: { readonly label: string; readonly value: number | null }) {
  return (
    <div className="rounded-lg bg-background/65 px-2.5 py-2">
      <dt className="text-2xs text-muted-foreground">{label}</dt>
      <dd className="mt-0.5 truncate text-xs font-medium tabular-nums text-foreground">
        {value === null ? "—" : formatWorkbookNumber(value)}
      </dd>
    </div>
  );
}

function ErrorNotice({ message }: { readonly message: string }) {
  return (
    <div
      aria-live="assertive"
      className="flex items-start gap-2 rounded-lg border border-destructive/30 bg-destructive/6 px-3 py-2.5 text-sm text-destructive-foreground"
    >
      <AlertCircleIcon className="mt-0.5 size-4 shrink-0" />
      <span>{message}</span>
    </div>
  );
}

function formatMetric(value: number | string): string {
  return typeof value === "number" ? formatWorkbookNumber(value) : value;
}

function formatFileSize(size: number): string {
  return size < 1024 * 1024
    ? `${formatWorkbookNumber(size / 1024)} KB`
    : `${formatWorkbookNumber(size / (1024 * 1024))} MB`;
}
