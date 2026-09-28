import type { PresentationChart, PresentationDeck } from "./presentationDeck";
import type { PresentationExportImage } from "./presentationExport";

type WorkRequest =
  | { type: "chart"; file: File }
  | {
      type: "export";
      deck: PresentationDeck;
      images: Readonly<Record<string, PresentationExportImage>>;
    };

function work<T>(request: WorkRequest): Promise<T> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./presentationWork.worker.ts", import.meta.url), {
      type: "module",
    });
    worker.onmessage = (
      event: MessageEvent<{ ok: true; result: T } | { ok: false; error: string }>,
    ) => {
      worker.terminate();
      if (event.data.ok) resolve(event.data.result);
      else reject(new Error(event.data.error));
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "Presentation worker stopped."));
    };
    worker.postMessage(request);
  });
}

export const importPresentationChart = (file: File) =>
  work<PresentationChart>({ type: "chart", file });
export const buildPresentation = (
  deck: PresentationDeck,
  images: Readonly<Record<string, PresentationExportImage>> = {},
) => work<string>({ type: "export", deck, images });
