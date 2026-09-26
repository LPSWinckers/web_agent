import { scopeProjectRef } from "@t3tools/client-runtime/environment";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";
import { FileSpreadsheetIcon, FileTextIcon, MessageSquareIcon, PlusIcon } from "lucide-react";
import { useCallback, useEffect, useMemo, useState } from "react";

import { SpreadsheetFileViewer } from "~/components/workbench/SpreadsheetFileViewer";
import { Button } from "~/components/ui/button";
import { Input } from "~/components/ui/input";
import { SidebarInset } from "~/components/ui/sidebar";
import { useNewThreadHandler } from "~/hooks/useHandleNewThread";
import { newProjectId } from "~/lib/utils";
import { useProjects } from "~/state/entities";
import { usePrimaryEnvironmentId } from "~/state/environments";
import { projectEnvironment } from "~/state/projects";
import { useEnvironmentQuery } from "~/state/query";
import { useAtomCommand } from "~/state/use-atom-command";
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
import { ConsultancyProjectFiles } from "./ConsultancyProjectFiles";

const MANIFEST_PATH = "consultancy/project.json";

function errorText(result: Parameters<typeof squashAtomCommandFailure>[0]): string {
  const error = squashAtomCommandFailure(result);
  return error instanceof Error ? error.message : "The request failed.";
}

export function ConsultancyWorkspace() {
  const projects = useProjects();
  const environmentId = usePrimaryEnvironmentId();
  const createProject = useAtomCommand(projectEnvironment.create, { reportFailure: false });
  const deleteProject = useAtomCommand(projectEnvironment.delete, { reportFailure: false });
  const writeFile = useAtomCommand(projectEnvironment.writeFile, { reportFailure: false });
  const newThread = useNewThreadHandler();
  const [selectedId, setSelectedId] = useState(readActiveConsultancyProject);
  const [customer, setCustomer] = useState(readPendingConsultancyCustomer);
  const [projectName, setProjectName] = useState("");
  const [folder, setFolder] = useState("");
  const [showCreate, setShowCreate] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [openedFile, setOpenedFile] = useState<File | null>(null);
  const [documentsOverride, setDocumentsOverride] = useState<ReadonlyArray<ProjectDocument> | null>(
    null,
  );
  const [customersVersion, setCustomersVersion] = useState(0);
  const project = projects.find((entry) => entry.id === selectedId) ?? null;
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

  useEffect(() => {
    const update = () => {
      setSelectedId(readActiveConsultancyProject());
      setDocumentsOverride(null);
      setOpenedFile(null);
    };
    window.addEventListener("consultancy-project-change", update);
    return () => window.removeEventListener("consultancy-project-change", update);
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
    if (!environmentId || !customer.trim() || !projectName.trim() || !folder.trim()) return;
    setBusy(true);
    setError("");
    try {
      const projectId = newProjectId();
      const result = await createProject({
        environmentId,
        input: {
          projectId,
          title: customerProjectTitle(customer, projectName),
          workspaceRoot: folder.trim(),
          createWorkspaceRootIfMissing: true,
          defaultModelSelection: null,
        },
      });
      if (result._tag === "Failure") throw new Error(errorText(result));
      selectConsultancyProject(projectId);
      setSelectedId(projectId);
      setShowCreate(false);
      setProjectName("");
      setFolder("");
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
    if (file) setOpenedFile(file);
    else setError("This workbook is stored in the browser where it was added.");
  };

  const openWorkbook = useCallback((file: File) => setOpenedFile(file), []);

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

  const startChat = async () => {
    if (!project) return;
    setError("");
    try {
      await newThread(scopeProjectRef(project.environmentId, project.id));
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not start a chat.");
    }
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

  if (openedFile)
    return (
      <SidebarInset className="h-dvh min-h-0 overflow-hidden">
        <SpreadsheetFileViewer file={openedFile} onClose={() => setOpenedFile(null)} inline />
      </SidebarInset>
    );

  return (
    <SidebarInset className="h-dvh min-h-0 overflow-y-auto">
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
              Choose a folder on the connected environment. Files in that folder will appear in this
              project.
            </p>
            <div className="mt-5 grid gap-3 sm:grid-cols-2">
              <Input
                list="consultancy-customers"
                placeholder="Customer name"
                value={customer}
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
                onChange={(event) => setProjectName(event.target.value)}
              />
              <Input
                className="sm:col-span-2"
                placeholder="Project folder, e.g. G:\\Consultancy\\Acme\\Market-study"
                value={folder}
                onChange={(event) => setFolder(event.target.value)}
              />
            </div>
            <div className="mt-4 flex gap-2">
              <Button
                disabled={
                  busy ||
                  !environmentId ||
                  !customer.trim() ||
                  !projectName.trim() ||
                  !folder.trim()
                }
                onClick={() => void create()}
              >
                Create project
              </Button>
              {project ? (
                <Button variant="ghost" onClick={() => setShowCreate(false)}>
                  Cancel
                </Button>
              ) : null}
            </div>
          </section>
        ) : null}
        {project && !showCreate ? (
          <>
            <ConsultancyProjectFiles
              key={`${project.environmentId}:${project.id}`}
              environmentId={project.environmentId}
              cwd={project.workspaceRoot}
              onOpenWorkbook={openWorkbook}
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
            <button
              type="button"
              onClick={() => void startChat()}
              className="mt-8 w-full rounded-xl border border-border bg-card p-5 text-left hover:border-primary/60"
            >
              <MessageSquareIcon className="size-5 text-primary" />
              <h2 className="mt-3 font-medium">Ask AI about this project</h2>
              <p className="mt-1 text-sm text-muted-foreground">
                Start a conversation about the files in this project.
              </p>
            </button>
            <div className="mt-8 border-t border-border pt-5">
              <Button variant="ghost" disabled={busy} onClick={() => void removeProject()}>
                Remove project
              </Button>
            </div>
          </>
        ) : null}
      </div>
    </SidebarInset>
  );
}
