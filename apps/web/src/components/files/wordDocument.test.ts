// @vitest-environment jsdom
import { DEFAULT_COMPANY_WORD_STANDARD } from "@t3tools/contracts";
import JSZip from "jszip";
import { describe, expect, it } from "vite-plus/test";
import {
  applyWordStructure,
  exportWordDocument,
  newWordDocument,
  openWordDocument,
  wordDocumentPath,
} from "./wordDocument";

const standard = {
  ...DEFAULT_COMPANY_WORD_STANDARD,
  fontFamily: "Aptos",
  headingColor: "123ABC",
  sections: ["Scope", "Recommendations"],
};

describe("Word documents", () => {
  it("uses one extension and keeps new documents in the Word folder", () => {
    expect(wordDocumentPath("Project plan.docx")).toBe("word/Project plan.docx");
  });

  it("round-trips a styled document with the company's sections", async () => {
    const content = newWordDocument("Plan <2026>", "Acme & Co", standard);
    const encoded = await exportWordDocument(content, standard);
    const zip = await JSZip.loadAsync(encoded, { base64: true });
    const styles = await zip.file("word/styles.xml")!.async("string");
    expect(styles).toContain("Aptos");
    expect(styles).toContain("123ABC");

    const bytes = Uint8Array.from(atob(encoded), (character) => character.charCodeAt(0));
    const opened = await openWordDocument(bytes.buffer);
    expect(opened.editable).toBe(true);
    expect(opened.content).toEqual(content);
    expect(await zip.file("word/document.xml")!.async("string")).toContain("Acme &amp; Co");
  });

  it("adds missing standard sections without duplicating existing headings", () => {
    const existing = newWordDocument("Plan", "Acme", { ...standard, sections: ["Scope"] });
    const updated = applyWordStructure(existing, standard);
    expect(updated.content?.filter((block) => block.type === "heading")).toHaveLength(3);
    expect(applyWordStructure(updated, standard)).toEqual(updated);
  });

  it("keeps bold and italic runs when a simple document is saved", async () => {
    const content = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [
            { type: "text", text: "Finding", marks: [{ type: "bold" }] },
            { type: "text", text: " and implication", marks: [{ type: "italic" }] },
          ],
        },
      ],
    };
    const base64 = await exportWordDocument(content, standard);
    const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
    expect((await openWordDocument(bytes.buffer)).content).toEqual(content);
  });

  it("exports line breaks created in the editor", async () => {
    const base64 = await exportWordDocument(
      {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              { type: "text", text: "First line" },
              { type: "hardBreak" },
              { type: "text", text: "Second line" },
            ],
          },
        ],
      },
      standard,
    );
    const zip = await JSZip.loadAsync(base64, { base64: true });
    expect(await zip.file("word/document.xml")!.async("string")).toContain("<w:br/>");
  });

  it("opens documents with tables read-only to avoid losing their contents", async () => {
    const encoded = await exportWordDocument(newWordDocument("Plan", "Acme", standard), standard);
    const zip = await JSZip.loadAsync(encoded, { base64: true });
    const source = await zip.file("word/document.xml")!.async("string");
    zip.file("word/document.xml", source.replace("<w:sectPr>", "<w:tbl/><w:sectPr>"));
    const bytes = await zip.generateAsync({ type: "arraybuffer" });
    const opened = await openWordDocument(bytes);
    expect(opened.editable).toBe(false);
    expect(opened.reason).toContain("Read-only");
  });
});
