import type { PresentationDeck } from "./presentationDeck";
import { importChart } from "./presentationChartImport";
import { exportPresentation } from "./presentationExport";

type Request = { type: "chart"; file: File } | { type: "export"; deck: PresentationDeck };

self.onmessage = (event: MessageEvent<Request>) => {
  void (async () => {
    try {
      const result =
        event.data.type === "chart"
          ? await importChart(event.data.file)
          : await exportPresentation(event.data.deck);
      self.postMessage({ ok: true, result });
    } catch (cause) {
      self.postMessage({
        ok: false,
        error: cause instanceof Error ? cause.message : "Presentation operation failed.",
      });
    }
  })();
};
