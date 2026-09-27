import type { JSONContent } from "@tiptap/core";
import type { CompanyWordStandard } from "@t3tools/contracts";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface OpenedWordDocument {
  readonly content: JSONContent;
  readonly editable: boolean;
  readonly reason?: string;
}

function xmlEscape(value: string): string {
  return value.replace(/[<>&"']/g, (character) => {
    switch (character) {
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case "&":
        return "&amp;";
      case '"':
        return "&quot;";
      default:
        return "&apos;";
    }
  });
}

function directChild(element: Element, name: string): Element | null {
  return [...element.children].find((child) => child.localName === name) ?? null;
}

function wordValue(element: Element | null): string | null {
  return element?.getAttributeNS(WORD_NAMESPACE, "val") ?? element?.getAttribute("w:val") ?? null;
}

function textNodes(paragraph: Element): JSONContent[] {
  const content: JSONContent[] = [];
  for (const run of paragraph.children) {
    if (run.localName !== "r") continue;
    const properties = directChild(run, "rPr");
    const marks = [
      ...(properties && directChild(properties, "b") ? [{ type: "bold" }] : []),
      ...(properties && directChild(properties, "i") ? [{ type: "italic" }] : []),
    ];
    let value = "";
    for (const item of run.children) {
      if (item.localName === "t") value += item.textContent ?? "";
      if (item.localName === "tab") value += "\t";
      if (item.localName === "br") value += "\n";
    }
    if (value) content.push({ type: "text", text: value, ...(marks.length ? { marks } : {}) });
  }
  return content;
}

function paragraphNode(paragraph: Element): JSONContent {
  const style = wordValue(directChild(directChild(paragraph, "pPr") ?? paragraph, "pStyle"));
  const heading = style === "Title" || style === "Heading1" || style === "Heading2";
  const flattenedText = [...paragraph.getElementsByTagNameNS(WORD_NAMESPACE, "t")]
    .map((text) => text.textContent ?? "")
    .join("");
  const content = isSimpleParagraph(paragraph)
    ? textNodes(paragraph)
    : flattenedText
      ? [{ type: "text", text: flattenedText }]
      : [];
  return {
    type: heading ? "heading" : "paragraph",
    ...(heading ? { attrs: { level: style === "Heading2" ? 2 : 1 } } : {}),
    content,
  };
}

function isSimpleParagraph(paragraph: Element): boolean {
  const properties = directChild(paragraph, "pPr");
  if (properties) {
    const style = wordValue(directChild(properties, "pStyle"));
    if (style && !["Normal", "Title", "Heading1", "Heading2"].includes(style)) return false;
    if ([...properties.children].some((child) => !["pStyle", "spacing"].includes(child.localName)))
      return false;
  }
  for (const child of paragraph.children) {
    if (child.localName === "pPr") continue;
    if (child.localName !== "r") return false;
    for (const item of child.children) {
      if (!["rPr", "t", "tab", "br"].includes(item.localName)) return false;
      if (
        item.localName === "rPr" &&
        [...item.children].some((mark) => !["b", "i"].includes(mark.localName))
      )
        return false;
    }
  }
  return true;
}

export async function openWordDocument(bytes: ArrayBuffer): Promise<OpenedWordDocument> {
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(bytes);
  const source = await zip.file("word/document.xml")?.async("string");
  if (!source) throw new Error("This Word document has no readable document body.");
  const xml = new DOMParser().parseFromString(source, "application/xml");
  if (xml.querySelector("parsererror")) throw new Error("This Word document is damaged.");
  const body = [...xml.getElementsByTagNameNS(WORD_NAMESPACE, "body")][0];
  if (!body) throw new Error("This Word document has no readable document body.");
  const paragraphs = [...xml.getElementsByTagNameNS(WORD_NAMESPACE, "p")];
  const content: JSONContent = {
    type: "doc",
    content: paragraphs.map(paragraphNode),
  };
  const unsupportedParts = Object.keys(zip.files).some((name) =>
    /^word\/(?:media\/|header\d*\.xml|footer\d*\.xml|footnotes\.xml|endnotes\.xml|comments\.xml)/i.test(
      name,
    ),
  );
  const unsupportedBody = [...body.children].some(
    (child) => !["p", "sectPr"].includes(child.localName),
  );
  const unsupportedParagraph = paragraphs.some((paragraph) => !isSimpleParagraph(paragraph));
  const editable = !(unsupportedParts || unsupportedBody || unsupportedParagraph);
  return {
    content,
    editable,
    ...(!editable
      ? {
          reason:
            "Read-only text preview. This document contains tables, images, or formatting the editor cannot preserve. Download the original for its full layout.",
        }
      : {}),
  };
}

export function newWordDocument(
  title: string,
  customer: string,
  standard: CompanyWordStandard,
): JSONContent {
  return {
    type: "doc",
    content: [
      { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: title }] },
      ...(customer
        ? [{ type: "paragraph", content: [{ type: "text", text: `Klant: ${customer}` }] }]
        : []),
      ...standard.sections.flatMap((section): JSONContent[] => [
        { type: "heading", attrs: { level: 2 }, content: [{ type: "text", text: section }] },
        { type: "paragraph", content: [] },
      ]),
    ],
  };
}

export function wordDocumentPath(title: string): string {
  const name = title
    .trim()
    .replace(/\.docx$/i, "")
    .replace(/[<>:"/\\|?*]/g, "-")
    .replace(/\p{Cc}/gu, "-")
    .slice(0, 90)
    .replace(/[. ]+$/, "");
  if (!name || name === "." || name === "..") throw new Error("Give the document a valid name.");
  return `word/${name}.docx`;
}

export function applyWordStructure(
  content: JSONContent,
  standard: CompanyWordStandard,
): JSONContent {
  const blocks = content.content ?? [];
  const headings = new Set(
    blocks
      .filter((block) => block.type === "heading")
      .map((block) =>
        (block.content ?? [])
          .map((item) => item.text ?? "")
          .join("")
          .trim()
          .toLowerCase(),
      ),
  );
  const missing = standard.sections.filter((section) => !headings.has(section.toLowerCase()));
  return {
    ...content,
    content: [
      ...blocks,
      ...missing.flatMap((section): JSONContent[] => [
        { type: "heading", attrs: { level: 2 }, content: [{ type: "text", text: section }] },
        { type: "paragraph", content: [] },
      ]),
    ],
  };
}

function runXml(node: JSONContent): string {
  if (node.type === "hardBreak") return "<w:r><w:br/></w:r>";
  if (node.type !== "text")
    throw new Error("This document contains content that cannot be exported to Word.");
  const marks = node.marks ?? [];
  if (marks.some((mark) => mark.type !== "bold" && mark.type !== "italic"))
    throw new Error("This document contains formatting that cannot be exported to Word.");
  const formatting = [
    ...(marks.some((mark) => mark.type === "bold") ? ["<w:b/>"] : []),
    ...(marks.some((mark) => mark.type === "italic") ? ["<w:i/>"] : []),
  ].join("");
  const pieces = (node.text ?? "").split("\n");
  const text = pieces
    .map((piece) => `<w:t xml:space="preserve">${xmlEscape(piece)}</w:t>`)
    .join("<w:br/>");
  return `<w:r>${formatting ? `<w:rPr>${formatting}</w:rPr>` : ""}${text}</w:r>`;
}

function blockXml(block: JSONContent): string {
  if (block.type !== "paragraph" && block.type !== "heading")
    throw new Error("This document contains blocks that cannot be exported to Word.");
  const style =
    block.type === "heading" ? (block.attrs?.level === 2 ? "Heading2" : "Heading1") : "Normal";
  const runs = (block.content ?? []).map(runXml).join("");
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${runs}</w:p>`;
}

function stylesXml(standard: CompanyWordStandard): string {
  const font = xmlEscape(standard.fontFamily);
  const style = (id: string, label: string, color: string, size: number, bold: boolean) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${label}"/><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:color w:val="${color}"/><w:sz w:val="${size}"/>${bold ? "<w:b/>" : ""}</w:rPr></w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${WORD_NAMESPACE}">${style("Normal", "Normal", standard.bodyColor, 22, false)}${style("Heading1", "heading 1", standard.headingColor, 36, true)}${style("Heading2", "heading 2", standard.accentColor, 28, true)}</w:styles>`;
}

export async function exportWordDocument(
  content: JSONContent,
  standard: CompanyWordStandard,
): Promise<string> {
  if (content.type !== "doc") throw new Error("This document cannot be exported to Word.");
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  );
  zip.file("word/styles.xml", stylesXml(standard));
  const body = (content.content ?? []).map(blockXml).join("");
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "base64", compression: "DEFLATE" });
}
