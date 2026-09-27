import type { EnvironmentProject } from "@t3tools/client-runtime/state/shell";

const CUSTOMERS_KEY = "t3-consultancy-customers";
const PROJECT_SEPARATOR = " / ";

export function isConsultancyInternalPath(path: string): boolean {
  return (
    path === "CONSULTANCY_CONTEXT.md" ||
    path === ".consultancy" ||
    path.startsWith(".consultancy/") ||
    path === "consultancy" ||
    path.startsWith("consultancy/") ||
    path === ".werkbestanden" ||
    path.startsWith(".werkbestanden/") ||
    path.split("/").at(-1) === ".keep" ||
    path.endsWith(".t3deck.json") ||
    path.endsWith(".t3chat.json")
  );
}

export interface ProjectDocument {
  readonly id: string;
  readonly name: string;
  readonly kind: "spreadsheet" | "document";
  readonly contextPath: string;
  readonly size: number;
}

export interface ProjectManifest {
  readonly version: 1;
  readonly documents: ReadonlyArray<ProjectDocument>;
}

export function customerProjectTitle(customer: string, project: string): string {
  return `${customer.trim()}${PROJECT_SEPARATOR}${project.trim()}`;
}

export function splitCustomerProject(project: EnvironmentProject) {
  const separator = project.title.indexOf(PROJECT_SEPARATOR);
  if (separator < 1) return null;
  const customer = project.title.slice(0, separator).trim();
  const name = project.title.slice(separator + PROJECT_SEPARATOR.length).trim();
  return customer && name ? { customer, name } : null;
}

export function readCustomers(): string[] {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(CUSTOMERS_KEY) ?? "[]");
    return Array.isArray(value)
      ? value.filter((name): name is string => typeof name === "string")
      : [];
  } catch {
    return [];
  }
}

export function saveCustomers(customers: ReadonlyArray<string>): void {
  localStorage.setItem(CUSTOMERS_KEY, JSON.stringify(customers));
}

export function parseManifest(contents: string | undefined): ProjectManifest {
  try {
    const value: unknown = JSON.parse(contents ?? "");
    if (
      typeof value === "object" &&
      value !== null &&
      "documents" in value &&
      Array.isArray(value.documents)
    ) {
      return {
        version: 1,
        documents: value.documents.filter(
          (document): document is ProjectDocument =>
            typeof document === "object" &&
            document !== null &&
            typeof document.id === "string" &&
            typeof document.name === "string" &&
            typeof document.contextPath === "string" &&
            (document.kind === "spreadsheet" || document.kind === "document") &&
            typeof document.size === "number",
        ),
      };
    }
  } catch {
    // A new project has no manifest yet.
  }
  return { version: 1, documents: [] };
}

function openFiles(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("t3-consultancy-files", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("files");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function saveOriginalFile(key: string, file: File): Promise<void> {
  const db = await openFiles();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction("files", "readwrite");
    transaction.objectStore("files").put(file, key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
  db.close();
}

export async function readOriginalFile(key: string): Promise<File | null> {
  const db = await openFiles();
  const file = await new Promise<File | null>((resolve, reject) => {
    const request = db.transaction("files", "readonly").objectStore("files").get(key);
    request.onsuccess = () => resolve(request.result instanceof File ? request.result : null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return file;
}

export async function deleteOriginalFile(key: string): Promise<void> {
  const db = await openFiles();
  await new Promise<void>((resolve, reject) => {
    const transaction = db.transaction("files", "readwrite");
    transaction.objectStore("files").delete(key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
  db.close();
}
