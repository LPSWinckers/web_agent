import { describe, expect, it } from "vite-plus/test";

import { fileChatSidecarPath, parseFileChatSidecar } from "./useFileChatThread";

describe("file chat sidecar", () => {
  it("keeps the thread association beside the exact file", () => {
    const path = "word/Brief.docx";
    expect(fileChatSidecarPath(path)).toBe("word/Brief.docx.t3chat.json");
    expect(fileChatSidecarPath("browser:workbook:123")).toBe(
      ".werkbestanden/browser%3Aworkbook%3A123.t3chat.json",
    );
    expect(
      parseFileChatSidecar('{"version":1,"path":"word/Brief.docx","threadId":"t-1"}', path),
    ).toBe("t-1");
    expect(
      parseFileChatSidecar('{"version":1,"path":"word/Other.docx","threadId":"t-2"}', path),
    ).toBeNull();
  });
});
