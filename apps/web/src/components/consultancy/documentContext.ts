import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.mjs?url";
import type { WorkbookResponse } from "../workbench/excelWorkbenchModel";

async function spreadsheetContext(file: File): Promise<string> {
  const worker = new Worker(new URL("../workbench/excelWorkbench.worker.ts", import.meta.url), {
    type: "module",
  });
  try {
    const request = (type: "load" | "context") =>
      new Promise<WorkbookResponse>((resolve, reject) => {
        worker.onmessage = (event: MessageEvent<WorkbookResponse>) => resolve(event.data);
        worker.onerror = (event) => reject(new Error(event.message));
        worker.postMessage(type === "load" ? { type, file, requestId: 1 } : { type, requestId: 2 });
      });
    const loaded = await request("load");
    if (loaded.type === "error") throw new Error(loaded.message);
    const context = await request("context");
    if (context.type === "error") throw new Error(context.message);
    if (context.type !== "context") throw new Error("Could not extract workbook context.");
    return context.text;
  } finally {
    worker.terminate();
  }
}

export async function extractDocumentContext(
  file: File,
): Promise<{ kind: "spreadsheet" | "document"; text: string }> {
  if (file.size > 20 * 1024 * 1024) throw new Error("Choose a file smaller than 20 MB.");
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "csv" || extension === "tsv") {
    return { kind: "spreadsheet", text: await spreadsheetContext(file) };
  }
  if (["txt", "md", "json", "xml", "html"].includes(extension ?? "")) {
    return { kind: "document", text: (await file.text()).slice(0, 800_000) };
  }
  if (extension === "docx") {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const xml = await zip.file("word/document.xml")?.async("string");
    if (!xml) throw new Error("This Word document could not be read.");
    const document = new DOMParser().parseFromString(xml, "application/xml");
    const paragraphs = [...document.getElementsByTagName("w:p")].map((paragraph) =>
      [...paragraph.getElementsByTagName("w:t")].map((text) => text.textContent ?? "").join(""),
    );
    const text = paragraphs.join("\n").trim();
    if (!text) throw new Error("This Word document has no readable text.");
    return { kind: "document", text: text.slice(0, 800_000) };
  }
  if (extension === "pdf") {
    const pdfjs = await import("pdfjs-dist");
    pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    const task = pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) });
    const pdf = await task.promise;
    try {
      const pages: string[] = [];
      for (let number = 1; number <= Math.min(pdf.numPages, 100); number++) {
        const page = await pdf.getPage(number);
        const content = await page.getTextContent();
        pages.push(
          `Page ${number}\n${content.items.map((item) => ("str" in item ? item.str : "")).join(" ")}`,
        );
        if (pages.join("\n").length > 800_000) break;
      }
      const text = pages.join("\n\n").trim();
      if (!text.replaceAll(/Page \d+/g, "").trim()) {
        throw new Error("This PDF has no selectable text. Convert it with OCR before adding it.");
      }
      return { kind: "document", text: text.slice(0, 800_000) };
    } finally {
      await task.destroy();
    }
  }
  throw new Error("Supported files: XLSX, CSV, TSV, PDF, DOCX, TXT, MD, JSON, XML, HTML.");
}
