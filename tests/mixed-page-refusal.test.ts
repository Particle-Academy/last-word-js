import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Agent, DocxReader, SchemaException, zipSync } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/*
 * Mixed page orientation, stage 1 (.ai/plans/mixed-page-orientation.md): a .docx
 * has one page setup until stage 2, so everything asking for more is refused
 * out loud, and everything a reader flattens is reported. Case for case with
 * PHP's MixedPageRefusalTest; the messages are the same bytes, the paths are
 * this engine's own form.
 */

const PAGE_REFUSED =
  "pageBreak.page is not implemented in last-word yet; remove page to keep one page setup, or write this document as a PDF with fine-print";
const SECTIONS_REFUSED = "there is no sections key; use pageBreak.page";
const NEAR_MISS_KEYS = [
  "size",
  "orientation",
  "margins",
  "header",
  "footer",
  "firstHeader",
  "firstFooter",
  "evenHeader",
  "evenFooter",
  "background",
  "border",
];

const para = (text = "Body") => ({ type: "paragraph", runs: [{ text }] });
const mixedPage = () => ({ orientation: "landscape", margins: { left: 28.8, right: 28.8 } });
const errorMap = (errors: { path: string; message: string }[]) =>
  Object.fromEntries(errors.map((e) => [e.path, e.message]));

const encoder = new TextEncoder();
function docx(body: string): Uint8Array {
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync([{ name: "word/document.xml", data: encoder.encode(xml) }]);
}

const BREAK = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`;
const LAST = `<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>`;
const TEXT = `<w:p><w:r><w:t>Text</w:t></w:r></w:p>`;

const fixture = (name: string) => new URL(`../test/fixtures/formats/${name}`, import.meta.url);

describe("validate", () => {
  it.each([
    ["a real page", { orientation: "landscape", margins: { left: 28.8 } }],
    ["an empty page", {}],
  ])("refuses pageBreak.page: %s", (_name, page) => {
    expect(Agent.validate({ blocks: [para(), { type: "pageBreak", page }, para()] })).toEqual([
      { path: "/blocks/1/page", message: PAGE_REFUSED },
    ]);
  });

  it("refuses pageBreak.page nested in a quote or a table cell too", () => {
    const pageBreak = { type: "pageBreak", page: mixedPage() };
    const errors = errorMap(
      Agent.validate({
        blocks: [
          { type: "quote", blocks: [pageBreak] },
          { type: "table", rows: [{ cells: [{ blocks: [pageBreak] }] }] },
        ],
      }),
    );
    expect(errors).toEqual({
      "/blocks/0/blocks/0/page": PAGE_REFUSED,
      "/blocks/1/rows/0/cells/0/blocks/0/page": PAGE_REFUSED,
    });
  });

  it.each(NEAR_MISS_KEYS)("refuses %s written straight on the break, naming pageBreak.page", (key) => {
    expect(Agent.validate({ blocks: [{ type: "pageBreak", [key]: "x" }] })).toEqual([
      { path: `/blocks/0/${key}`, message: `a pageBreak has no ${key} key; use pageBreak.page` },
    ]);
  });

  it("stays lenient about any other extra key on a break: agents decorate", () => {
    expect(Agent.validate({ blocks: [{ type: "pageBreak", note: "mine", id: "b1" }] })).toEqual([]);
  });

  it.each([
    ["a list", [{ page: { orientation: "landscape" } }]],
    ["empty", []],
  ])("refuses a top-level sections key: %s", (_name, sections) => {
    expect(Agent.validate({ sections, blocks: [] })).toEqual([{ path: "/sections", message: SECTIONS_REFUSED }]);
  });

  it("still says sections is wrong when blocks is missing as well", () => {
    expect(Object.keys(errorMap(Agent.validate({ sections: [] })))).toContain("/sections");
  });
});

describe("repair: the trap", () => {
  it("keeps page, so an unrelated repair cannot launder it into ok: true", () => {
    const result: Any = Agent.validateAndRepair({
      blocks: ["needs repair", { type: "pageBreak", page: mixedPage() }, para()],
    });
    expect(result.ok).toBe(false);
    expect(result.schema.blocks[0]).toEqual(para("needs repair"));
    expect(result.schema.blocks[1]).toEqual({ type: "pageBreak", page: mixedPage() });
    expect(errorMap(result.errors)["/blocks/1/page"]).toBe(PAGE_REFUSED);
  });

  it("keeps an empty page as well: {} is still a request", () => {
    const result: Any = Agent.validateAndRepair({ blocks: ["needs repair", { type: "pageBreak", page: {} }] });
    expect(result.ok).toBe(false);
    expect(result.schema.blocks[1]).toEqual({ type: "pageBreak", page: {} });
  });

  it("keeps page nested in a quote and a cell", () => {
    const pageBreak = { type: "pageBreak", page: mixedPage() };
    const result: Any = Agent.validateAndRepair({
      blocks: [
        "needs repair",
        { type: "quote", blocks: [pageBreak] },
        { type: "table", rows: [{ cells: [{ blocks: [pageBreak] }] }] },
      ],
    });
    expect(result.ok).toBe(false);
    expect(result.schema.blocks[1].blocks[0]).toEqual(pageBreak);
    expect(result.schema.blocks[2].rows[0].cells[0].blocks[0]).toEqual(pageBreak);
  });

  it("neither strips nor moves a near-miss key: it stays refused until stage 2", () => {
    const result: Any = Agent.validateAndRepair({
      blocks: ["needs repair", { type: "pageBreak", orientation: "landscape", margins: { left: 10 } }],
    });
    expect(result.ok).toBe(false);
    expect(result.schema.blocks[1]).toEqual({ type: "pageBreak", orientation: "landscape", margins: { left: 10 } });
    const errors = errorMap(result.errors);
    expect(errors["/blocks/1/orientation"]).toBe("a pageBreak has no orientation key; use pageBreak.page");
    expect(errors["/blocks/1/margins"]).toBe("a pageBreak has no margins key; use pageBreak.page");
  });

  it("keeps a top-level sections key and fails", () => {
    const result: Any = Agent.validateAndRepair({ sections: [{ x: 1 }], blocks: ["needs repair"] });
    expect(result.ok).toBe(false);
    expect(result.schema.sections).toEqual([{ x: 1 }]);
    expect(errorMap(result.errors)["/sections"]).toBe(SECTIONS_REFUSED);
  });
});

describe("writing", () => {
  it.each([
    ["page", { blocks: [{ type: "pageBreak", page: { orientation: "landscape" } }] }, "/blocks/0/page", PAGE_REFUSED],
    [
      "near miss",
      { blocks: [{ type: "pageBreak", orientation: "landscape" }] },
      "/blocks/0/orientation",
      "a pageBreak has no orientation key; use pageBreak.page",
    ],
    ["sections", { sections: [], blocks: [] }, "/sections", SECTIONS_REFUSED],
  ])("toBytes refuses %s", (_name, doc, path, message) => {
    let thrown: unknown;
    try {
      Agent.toBytes(doc);
    } catch (e) {
      thrown = e;
    }
    expect(thrown).toBeInstanceOf(SchemaException);
    expect(errorMap((thrown as SchemaException).errors)).toEqual({ [path]: message });
  });

  it("writes a plain break exactly as before", () => {
    const doc = { blocks: [para("one"), { type: "pageBreak" }, para("two")] };
    expect(Agent.validate(doc)).toEqual([]);
    expect((Agent.read(Agent.toBytes(doc)) as Any).blocks).toEqual(doc.blocks);
  });
});

describe("reading", () => {
  it("reads a real two-section file as PHP does, report included", () => {
    const doc: Any = Agent.read(new Uint8Array(readFileSync(fixture("two-sections.docx"))));
    const php = JSON.parse(readFileSync(fixture("two-sections.read.json"), "utf8"));

    expect(doc.unrepresented).toEqual([{ where: "/page", kind: "layout-not-preserved", count: 1 }]);
    expect(doc).toEqual(php);
  });

  it.each([
    ["one section", TEXT + LAST, null],
    ["no sectPr at all", TEXT, null],
    ["two sections", TEXT + BREAK + TEXT + LAST, 1],
    ["three sections", TEXT + BREAK + TEXT + BREAK + TEXT + LAST, 2],
    ["a break and no final sectPr", TEXT + BREAK + TEXT, 1],
    ["the break on a paragraph with text", `<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>` + TEXT + LAST, 1],
    ["inside an sdt", `<w:sdt><w:sdtContent>${BREAK}</w:sdtContent></w:sdt>` + TEXT + LAST, 1],
    ["inside customXml", `<w:customXml>${BREAK}</w:customXml>` + TEXT + LAST, 1],
    ["a tracked earlier sectPr is not a section", TEXT + `<w:sectPr><w:sectPrChange><w:sectPr/></w:sectPrChange></w:sectPr>`, null],
  ])("counts the section breaks the read flattened: %s", (_name, body, count) => {
    const doc: Any = new DocxReader().read(docx(body));
    if (count === null) {
      expect(doc).not.toHaveProperty("unrepresented");
    } else {
      expect(doc.unrepresented).toEqual([{ where: "/page", kind: "layout-not-preserved", count }]);
    }
  });

  it("reads the same blocks with or without the section break", () => {
    const onText = `<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>`;
    const { unrepresented: _dropped, ...withBreak }: Any = Agent.read(docx(onText + TEXT + LAST));
    expect(withBreak).toEqual(Agent.read(docx(TEXT + TEXT + LAST)));
  });

  it("keeps today's empty paragraph for a section break on an empty one (PHP and Python drop it)", () => {
    const doc: Any = Agent.read(docx(TEXT + BREAK + TEXT + LAST));
    expect(doc.blocks).toEqual([para("Text"), { type: "paragraph", runs: [] }, para("Text")]);
  });

  it("never reports a document the writer wrote", () => {
    const doc = Agent.read(
      Agent.toBytes({
        page: { orientation: "landscape", header: { left: "H" } },
        blocks: [para(), { type: "pageBreak" }, para()],
      }),
    );
    expect(doc).not.toHaveProperty("unrepresented");
  });

  it("describe says so, in PHP's words", () => {
    const note = (count: number) => [{ where: "/page", kind: "layout-not-preserved", count }];
    expect(Agent.describe({ blocks: [], unrepresented: note(2) })).toContain(
      "Layout not preserved: 2 section breaks flattened into the last section's page setup",
    );
    expect(Agent.describe({ blocks: [], unrepresented: note(1) })).toContain(
      "Layout not preserved: 1 section break flattened into the last section's page setup",
    );
    expect(Agent.describe({ blocks: [] })).not.toContain("Layout not preserved");
  });
});
