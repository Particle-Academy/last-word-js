import { describe, it, expect } from "vitest";
import { Agent, zipSync } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const encoder = new TextEncoder();

/** Build a docx zip from a raw word/document.xml body (Word-tolerance harness). */
function docxFromBody(body: string, extraParts: { name: string; xml: string }[] = []): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync([
    { name: "word/document.xml", data: encoder.encode(documentXml) },
    ...extraParts.map((p) => ({ name: p.name, data: encoder.encode(p.xml) })),
  ]);
}

describe("Word-file tolerance (vector 5)", () => {
  it("reads outlineLvl headings and unknown elements without throwing", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Chapter</w:t></w:r></w:p>` +
        `<w:customXml><w:p><w:r><w:t>inside an unknown wrapper</w:t></w:r></w:p></w:customXml>` +
        `<w:frobnicator w:mystery="1"/>` +
        `<w:p><w:r><w:t>plain paragraph</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([
      { type: "heading", level: 1, runs: [{ text: "Chapter" }] },
      { type: "paragraph", runs: [{ text: "inside an unknown wrapper" }] },
      { type: "paragraph", runs: [{ text: "plain paragraph" }] },
    ]);
  });

  it("clamps outlineLvl beyond 5 into heading level 6", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:outlineLvl w:val="8"/></w:pPr><w:r><w:t>Deep</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0]).toEqual({ type: "heading", level: 6, runs: [{ text: "Deep" }] });
  });

  it("reads pStyle Heading1-9 headings (Word style ids)", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Heading9"/></w:pPr><w:r><w:t>Nine</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0].level).toBe(2);
    expect(doc.blocks[1].level).toBe(6);
  });

  it("buckets unknown numIds as unordered lists (no numbering.xml)", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>` +
        `<w:r><w:t>item</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0]).toEqual({ type: "list", items: [{ runs: [{ text: "item" }] }] });
  });

  it("maps decimal numFmt to ordered via numbering.xml", () => {
    const numbering =
      `<?xml version="1.0"?>` +
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="9"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="42"><w:abstractNumId w:val="9"/></w:num></w:numbering>`;
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="42"/></w:numPr></w:pPr>` +
        `<w:r><w:t>first</w:t></w:r></w:p>`,
      [{ name: "word/numbering.xml", xml: numbering }],
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0].ordered).toBe(true);
  });

  it("maps named w:highlight values to nearest hex", () => {
    const bytes = docxFromBody(
      `<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>hot</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0].runs[0]).toEqual({ text: "hot", highlight: "#FFFF00" });
  });

  it("reads mono rFonts as inline code (no rStyle needed)", () => {
    const bytes = docxFromBody(
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>x = 1</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0].runs[0]).toEqual({ text: "x = 1", code: true });
  });

  it("treats bottom-border-only empty paragraphs as hr and w:br page as pageBreak", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>` +
        `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([{ type: "hr" }, { type: "pageBreak" }]);
  });

  it("groups consecutive Quote-styled paragraphs into one quote block", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>first</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:t>second</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([
      {
        type: "quote",
        blocks: [
          { type: "paragraph", runs: [{ text: "first" }] },
          { type: "paragraph", runs: [{ text: "second" }] },
        ],
      },
    ]);
  });

  it("groups consecutive CodeBlock-styled paragraphs into one code block", () => {
    const bytes = docxFromBody(
      `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr><w:r><w:t>line 1</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr><w:r><w:t>line 2</w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([{ type: "code", text: "line 1\nline 2" }]);
  });

  it("reads unknown SDT wrappers transparently", () => {
    const bytes = docxFromBody(
      `<w:sdt><w:sdtPr><w:tag w:val="someVendorTag"/></w:sdtPr>` +
        `<w:sdtContent><w:p><w:r><w:t>wrapped</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([{ type: "paragraph", runs: [{ text: "wrapped" }] }]);
  });

  it("never throws on tabs, soft breaks, and empty runs", () => {
    const bytes = docxFromBody(
      `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r><w:r><w:t></w:t></w:r></w:p>`,
    );
    const doc: Any = Agent.read(bytes);
    expect(doc.blocks[0].runs).toEqual([{ text: "a\tb\nc" }]);
  });
});
