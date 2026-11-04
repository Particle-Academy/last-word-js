import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Agent, DocxReader, SchemaException, zipSync } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/*
 * Mixed page orientation (.ai/plans/mixed-page-orientation.md): what is still
 * refused now that `pageBreak.page` is honoured (see mixed-pages.test.ts). Page keys
 * written straight on a break, and a top-level `sections`. Case for case with PHP's
 * MixedPageRefusalTest; the messages are the same bytes, the paths are this engine's
 * own form.
 */

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

describe("repair", () => {
  it("keeps a top-level sections key and fails", () => {
    const result: Any = Agent.validateAndRepair({ sections: [{ x: 1 }], blocks: ["needs repair"] });
    expect(result.ok).toBe(false);
    expect(result.schema.sections).toEqual([{ x: 1 }]);
    expect(errorMap(result.errors)["/sections"]).toBe(SECTIONS_REFUSED);
  });
});

describe("writing", () => {
  it.each([
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
