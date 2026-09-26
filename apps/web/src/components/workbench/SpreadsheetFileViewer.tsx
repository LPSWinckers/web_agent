import {
  ChevronLeftIcon,
  ChevronRightIcon,
  FileSpreadsheetIcon,
  MessageSquareIcon,
  SendIcon,
  XIcon,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState } from "react";

import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import {
  formatWorkbookCell,
  formatWorkbookNumber,
  WORKBOOK_PAGE_SIZE,
  type WorkbookCell,
  type WorkbookCommand,
  type WorkbookResponse,
  type WorkbookSheetOverview,
} from "./excelWorkbenchModel";

interface PendingRequest {
  readonly resolve: (response: Exclude<WorkbookResponse, { type: "error" }>) => void;
  readonly reject: (error: Error) => void;
}

interface ChatMessage {
  readonly id: number;
  readonly role: "user" | "answer";
  readonly text: string;
  readonly sheet: string;
}

export function SpreadsheetFileViewer({
  file,
  onClose,
  inline = false,
}: {
  file: File;
  onClose?: () => void;
  inline?: boolean;
}) {
  const workerRef = useRef<Worker | null>(null);
  const pendingRef = useRef(new Map<number, PendingRequest>());
  const requestIdRef = useRef(0);
  const queryIdRef = useRef(0);
  const messageIdRef = useRef(0);
  const [sheetNames, setSheetNames] = useState<ReadonlyArray<string>>([]);
  const [overview, setOverview] = useState<WorkbookSheetOverview | null>(null);
  const [rows, setRows] = useState<ReadonlyArray<ReadonlyArray<WorkbookCell>>>([]);
  const [filteredRowCount, setFilteredRowCount] = useState(0);
  const [page, setPage] = useState(0);
  const [search, setSearch] = useState("");
  const [chatOpen, setChatOpen] = useState(false);
  const [question, setQuestion] = useState("");
  const [messages, setMessages] = useState<ReadonlyArray<ChatMessage>>([]);
  const [busy, setBusy] = useState(true);
  const [asking, setAsking] = useState(false);
  const [error, setError] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const sendCommand = useCallback((command: WorkbookCommand) => {
    return new Promise<Exclude<WorkbookResponse, { type: "error" }>>((resolve, reject) => {
      const worker = workerRef.current;
      if (!worker) {
        reject(new Error("The spreadsheet is not ready."));
        return;
      }
      const requestId = ++requestIdRef.current;
      pendingRef.current.set(requestId, { resolve, reject });
      worker.postMessage({ ...command, requestId });
    });
  }, []);

  useEffect(() => {
    let disposed = false;
    const worker = new Worker(new URL("./excelWorkbench.worker.ts", import.meta.url), {
      type: "module",
    });
    workerRef.current = worker;
    worker.onmessage = (event: MessageEvent<WorkbookResponse>) => {
      const response = event.data;
      const pending = pendingRef.current.get(response.requestId);
      if (!pending) return;
      pendingRef.current.delete(response.requestId);
      if (response.type === "error") pending.reject(new Error(response.message));
      else pending.resolve(response);
    };
    worker.onerror = (event) => {
      const failure = new Error(event.message || "The spreadsheet reader stopped.");
      for (const pending of pendingRef.current.values()) pending.reject(failure);
      pendingRef.current.clear();
      setError(failure.message);
      setBusy(false);
    };
    void sendCommand({ type: "load", file })
      .then((response) => {
        if (disposed) return;
        if (response.type !== "loaded") return;
        setSheetNames(response.sheetNames);
        setOverview(response.overview);
        setRows(response.rows);
        setFilteredRowCount(response.filteredRowCount);
      })
      .catch((cause: unknown) => {
        if (!disposed)
          setError(cause instanceof Error ? cause.message : "Could not open this spreadsheet.");
      })
      .finally(() => {
        if (!disposed) setBusy(false);
      });
    return () => {
      disposed = true;
      queryIdRef.current++;
      workerRef.current = null;
      worker.terminate();
      for (const pending of pendingRef.current.values()) {
        pending.reject(new Error("The file viewer was closed."));
      }
      pendingRef.current.clear();
    };
  }, [file, sendCommand]);

  const activeSheetName = overview?.name;
  useEffect(() => {
    if (!activeSheetName) return;
    const queryId = ++queryIdRef.current;
    const timeout = window.setTimeout(
      () => {
        void sendCommand({ type: "query", term: search, page })
          .then((response) => {
            if (queryId !== queryIdRef.current || response.type !== "rows") return;
            setRows(response.rows);
            setFilteredRowCount(response.filteredRowCount);
          })
          .catch((cause: unknown) => {
            if (queryId === queryIdRef.current) {
              setError(cause instanceof Error ? cause.message : "Could not search this sheet.");
            }
          });
      },
      search ? 180 : 0,
    );
    return () => window.clearTimeout(timeout);
  }, [activeSheetName, page, search, sendCommand]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ block: "end" });
  }, [messages, asking]);

  const selectSheet = async (name: string) => {
    if (name === overview?.name || busy || asking) return;
    setBusy(true);
    setError("");
    queryIdRef.current++;
    try {
      const response = await sendCommand({ type: "sheet", name });
      if (response.type !== "sheet") return;
      setOverview(response.overview);
      setRows(response.rows);
      setFilteredRowCount(response.filteredRowCount);
      setPage(0);
      setSearch("");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not open this sheet.");
    } finally {
      setBusy(false);
    }
  };

  const ask = async () => {
    const text = question.trim();
    if (!text || !overview || asking || busy) return;
    const sheet = overview.name;
    setMessages((current) => [
      ...current,
      { id: ++messageIdRef.current, role: "user", text, sheet },
    ]);
    setQuestion("");
    setAsking(true);
    try {
      const response = await sendCommand({ type: "ask", question: text });
      if (response.type === "answer") {
        setMessages((current) => [
          ...current,
          { id: ++messageIdRef.current, role: "answer", text: response.text, sheet },
        ]);
      }
    } catch (cause) {
      setMessages((current) => [
        ...current,
        {
          id: ++messageIdRef.current,
          role: "answer",
          text: cause instanceof Error ? cause.message : "Could not answer that question.",
          sheet,
        },
      ]);
    } finally {
      setAsking(false);
    }
  };

  const totalPages = Math.max(1, Math.ceil(filteredRowCount / WORKBOOK_PAGE_SIZE));

  return (
    <main
      className={`${inline ? "flex min-h-0 flex-1 flex-col" : "fixed inset-0 z-60 flex min-h-0 flex-col"} bg-background text-foreground`}
    >
      <header className="flex h-14 shrink-0 items-center gap-3 border-b border-border px-3 sm:px-5">
        {onClose ? (
          <Button aria-label="Back to project" onClick={onClose} size="icon-sm" variant="ghost">
            <ChevronLeftIcon />
          </Button>
        ) : null}
        <FileSpreadsheetIcon className="size-5 shrink-0 text-primary" />
        <div className="min-w-0 flex-1">
          <h1 className="truncate text-sm font-semibold">{file.name}</h1>
          <p className="truncate text-xs text-muted-foreground">
            {overview
              ? `${overview.name} · ${formatWorkbookNumber(overview.rowCount)} rows`
              : "Opening file"}
          </p>
        </div>
        <Button
          aria-expanded={chatOpen}
          aria-label={chatOpen ? "Close file chat" : "Open file chat"}
          onClick={() => setChatOpen((current) => !current)}
          size="sm"
          variant={chatOpen ? "secondary" : "outline"}
        >
          <MessageSquareIcon />
          <span className="hidden sm:inline">Ask about file</span>
        </Button>
      </header>

      <div className="relative flex min-h-0 flex-1">
        <section aria-label="Spreadsheet viewer" className="flex min-w-0 flex-1 flex-col">
          {error ? (
            <p
              role="alert"
              className="border-b border-destructive/30 bg-destructive/10 px-4 py-2 text-sm text-destructive"
            >
              {error}
            </p>
          ) : null}
          {overview ? (
            <>
              <div className="flex shrink-0 flex-wrap items-center justify-between gap-2 border-b border-border px-3 py-2">
                <nav aria-label="Worksheets" className="flex min-w-0 flex-1 gap-1 overflow-x-auto">
                  {sheetNames.map((name) => (
                    <button
                      aria-current={name === overview.name ? "page" : undefined}
                      className={`shrink-0 rounded-md px-3 py-1.5 text-sm ${name === overview.name ? "bg-accent font-medium text-foreground" : "text-muted-foreground hover:bg-accent/60"}`}
                      disabled={busy || asking}
                      key={name}
                      onClick={() => void selectSheet(name)}
                      type="button"
                    >
                      {name}
                    </button>
                  ))}
                </nav>
                <Input
                  aria-label="Search spreadsheet rows"
                  className="w-full sm:w-56"
                  disabled={busy}
                  onChange={(event) => {
                    setSearch(event.currentTarget.value);
                    setPage(0);
                  }}
                  placeholder="Search rows"
                  value={search}
                />
              </div>
              <div className="min-h-0 flex-1 overflow-auto">
                <table className="w-full min-w-max border-collapse text-xs">
                  <thead className="sticky top-0 z-10 bg-muted">
                    <tr>
                      <th className="sticky left-0 z-20 border-b border-r border-border bg-muted px-3 py-2 text-right font-medium">
                        #
                      </th>
                      {overview.columns.map((column) => (
                        <th
                          className="max-w-72 border-b border-r border-border px-3 py-2 text-left font-medium"
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
                        className="hover:bg-accent/30"
                        key={`${overview.name}-${page}-${rowIndex}`}
                      >
                        <td className="sticky left-0 border-b border-r border-border bg-background px-3 py-2 text-right tabular-nums text-muted-foreground">
                          {formatWorkbookNumber(page * WORKBOOK_PAGE_SIZE + rowIndex + 1)}
                        </td>
                        {overview.columns.map((column, columnIndex) => (
                          <td
                            className="max-w-72 border-b border-r border-border px-3 py-2"
                            key={column.name}
                          >
                            <span className="block truncate">
                              {formatWorkbookCell(row[columnIndex] ?? null) || " "}
                            </span>
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {rows.length === 0 ? (
                  <p className="p-8 text-center text-sm text-muted-foreground">
                    {search ? "No rows match your search." : "This sheet has no data rows."}
                  </p>
                ) : null}
              </div>
              <footer className="flex h-11 shrink-0 items-center justify-between gap-3 border-t border-border px-3 text-xs text-muted-foreground">
                <span>
                  {formatWorkbookNumber(filteredRowCount)} {search ? "matching" : "total"} rows
                </span>
                <div className="flex items-center gap-2">
                  <span>
                    Page {page + 1} of {totalPages}
                  </span>
                  <Button
                    aria-label="Previous page"
                    disabled={page === 0 || busy}
                    onClick={() => setPage((current) => current - 1)}
                    size="icon-xs"
                    variant="ghost"
                  >
                    <ChevronLeftIcon />
                  </Button>
                  <Button
                    aria-label="Next page"
                    disabled={page + 1 >= totalPages || busy}
                    onClick={() => setPage((current) => current + 1)}
                    size="icon-xs"
                    variant="ghost"
                  >
                    <ChevronRightIcon />
                  </Button>
                </div>
              </footer>
            </>
          ) : (
            <div className="flex flex-1 items-center justify-center px-6 text-sm text-muted-foreground">
              {busy ? "Opening spreadsheet…" : "This spreadsheet could not be displayed."}
            </div>
          )}
        </section>

        {chatOpen ? (
          <aside
            aria-label="File chat"
            className="absolute inset-0 z-20 flex flex-col border-l border-border bg-background sm:static sm:w-88 sm:shrink-0"
          >
            <div className="flex h-12 shrink-0 items-center justify-between border-b border-border px-4">
              <div>
                <h2 className="text-sm font-semibold">Ask about this sheet</h2>
                <p className="text-xs text-muted-foreground">{overview?.name ?? file.name}</p>
              </div>
              <Button
                aria-label="Close file chat"
                onClick={() => setChatOpen(false)}
                size="icon-xs"
                variant="ghost"
              >
                <XIcon />
              </Button>
            </div>
            <div aria-live="polite" className="min-h-0 flex-1 space-y-3 overflow-y-auto p-4">
              {messages.length === 0 ? (
                <p className="text-sm leading-6 text-muted-foreground">
                  Ask for row counts, missing values, duplicates, or the total or average of a
                  numeric column. Answers use the selected worksheet.
                </p>
              ) : null}
              {messages.map((message) => (
                <div
                  className={`rounded-lg px-3 py-2 text-sm leading-5 ${message.role === "user" ? "ml-6 bg-primary text-primary-foreground" : "mr-6 bg-muted"}`}
                  key={message.id}
                >
                  {message.role === "answer" ? (
                    <p className="mb-1 text-xs opacity-65">{message.sheet}</p>
                  ) : null}
                  <p className="whitespace-pre-wrap">{message.text}</p>
                </div>
              ))}
              {asking ? <p className="text-xs text-muted-foreground">Checking worksheet…</p> : null}
              <div ref={messagesEndRef} />
            </div>
            <form
              className="flex shrink-0 gap-2 border-t border-border p-3"
              onSubmit={(event) => {
                event.preventDefault();
                void ask();
              }}
            >
              <Input
                aria-label="Question about spreadsheet"
                disabled={!overview || busy || asking}
                onChange={(event) => setQuestion(event.currentTarget.value)}
                placeholder="Ask about this sheet"
                value={question}
              />
              <Button
                aria-label="Send question"
                disabled={!question.trim() || !overview || busy || asking}
                size="icon-sm"
              >
                <SendIcon />
              </Button>
            </form>
          </aside>
        ) : null}
      </div>
    </main>
  );
}
