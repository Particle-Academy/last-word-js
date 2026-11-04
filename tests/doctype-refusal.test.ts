import { readFileSync, writeFileSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { Agent, unzipSync, zipSync } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/*
 * Every XML part of an uploaded file is read through one gate, which refuses a
 * DOCTYPE. Through the public reader, because what matters is what a hostile upload
 * gets back. The same fixtures as PHP's PartXmlTest and Python's test_part_xml.
 *
 * MARKER is the visible text of each part in a clean file, and ENTITY the value its
 * replacement declares, so ENTITY in the result means it expanded.
 */
const MARKER = "LWBASE";
const ENTITY = "LWINT";
const DOCTYPE = `<!DOCTYPE r [<!ENTITY x "${ENTITY}">]>`;
const enc = new TextEncoder();
const dec = new TextDecoder();

function docx(): Record<string, string | Uint8Array> {
  const bytes = Agent.toBytes({
    title: MARKER,
    defaultFont: MARKER,
    page: { header: { left: MARKER } },
    blocks: [
      {
        type: "paragraph",
        runs: [{ text: MARKER }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: MARKER }] }] } }],
      },
      {
        type: "chart",
        option: { title: { text: MARKER }, xAxis: { type: "category", data: ["a"] }, yAxis: { type: "value" }, series: [{ type: "bar", data: [1] }] },
      },
    ],
  } as Any);
  return unzipSync(bytes) as Record<string, Uint8Array>;
}

const text = (part: string | Uint8Array): string => (typeof part === "string" ? part : dec.decode(part));

function zip(parts: Record<string, string | Uint8Array>): Uint8Array {
  return zipSync(Object.entries(parts).map(([name, body]) => ({ name, data: typeof body === "string" ? enc.encode(body) : body })));
}

function withDoctype(xml: string, doctype = DOCTYPE): string {
  const m = /^<\?xml[^>]*\?>\s*/.exec(xml);
  return m ? m[0] + doctype + xml.slice(m[0].length) : doctype + xml;
}

const declare = (xml: string, encoding: string): string => xml.replace(/^<\?xml[^>]*\?>/, `<?xml version="1.0" encoding="${encoding}"?>`);
const utf16le = (s: string): Uint8Array => new Uint8Array(Buffer.from(s, "utf16le"));
const utf16be = (s: string): Uint8Array => new Uint8Array(Buffer.from(s, "utf16le").swap16());
const concat = (a: number[], b: Uint8Array): Uint8Array => new Uint8Array([...a, ...b]);

const ENCODINGS: Record<string, (x: string) => string | Uint8Array> = {
  "plain UTF-8": (x) => x,
  "UTF-8 with a BOM": (x) => concat([0xef, 0xbb, 0xbf], enc.encode(x)),
  "UTF-16LE with a BOM": (x) => concat([0xff, 0xfe], utf16le(declare(x, "UTF-16"))),
  "UTF-16BE with a BOM": (x) => concat([0xfe, 0xff], utf16be(declare(x, "UTF-16"))),
  "UTF-16LE without a BOM": (x) => utf16le(declare(x, "UTF-16")),
  "past a long comment": (x) => x.replace(/^(<\?xml[^>]*\?>\s*)/, `$1<!-- ${"p".repeat(70000)} -->`),
};
const PARTS = ["word/document.xml", "word/header1.xml", "word/footnotes.xml", "docProps/core.xml", "word/styles.xml", "word/charts/chart1.xml"];
const CASES = PARTS.flatMap((part) => Object.keys(ENCODINGS).map((encoding) => [part, encoding] as const));

function read(parts: Record<string, string | Uint8Array>): string {
  try {
    return JSON.stringify(Agent.read(zip(parts)));
  } catch {
    return "(refused)";
  }
}

const count = (s: string): number => s.split(MARKER).length - 1;
const CLEAN = count(read(docx()));

describe("every XML part goes through the gate", () => {
  it("reads the marker from every part of a clean file, so a missing one means refused", () => {
    expect(CLEAN).toBe(6);
  });

  it.each(CASES)("%s, %s: expands no entity", (part, encoding) => {
    const parts = docx();
    parts[part] = ENCODINGS[encoding]!(withDoctype(text(parts[part]!).split(MARKER).join("&x;")));
    expect(read(parts)).not.toContain(ENTITY);
  });

  it.each(CASES)("%s, %s: takes nothing from a part declaring a DOCTYPE", (part, encoding) => {
    const parts = docx();
    parts[part] = ENCODINGS[encoding]!(withDoctype(text(parts[part]!)));
    expect(count(read(parts))).toBeLessThan(CLEAN);
  });

  it.each(PARTS)("%s: still reads a clean part in UTF-8, with a BOM or past a comment", (part) => {
    for (const encoding of ["plain UTF-8", "UTF-8 with a BOM", "past a long comment"]) {
      const parts = docx();
      parts[part] = ENCODINGS[encoding]!(text(parts[part]!));
      expect(count(read(parts))).toBe(CLEAN);
    }
  });

  it("does not take markup out of a DOCTYPE's internal subset", () => {
    const parts = docx();
    const body = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>real</w:t></w:r></w:p></w:body></w:document>`;
    parts["word/document.xml"] = `<!DOCTYPE r [<!ENTITY x "a><w:document><w:body><w:p><w:r><w:t>${ENTITY}</w:t></w:r></w:p></w:body></w:document>">]>${body}`;
    expect(read(parts)).not.toContain(ENTITY);
  });

  it("treats a refused body part as an unparsable one: no throw, no blocks", () => {
    const refused = docx();
    refused["word/document.xml"] = withDoctype(text(refused["word/document.xml"]!));
    const unparsable = docx();
    unparsable["word/document.xml"] = "";
    expect((JSON.parse(read(refused)) as Any).blocks).toEqual((JSON.parse(read(unparsable)) as Any).blocks);
  });

  it("follows no relationship whose target comes from an entity", () => {
    const parts = docx();
    const rels = text(parts["word/_rels/document.xml.rels"]!).replace('Target="header1.xml"', 'Target="&x;"');
    parts["word/_rels/document.xml.rels"] = withDoctype(rels, `<!DOCTYPE r [<!ENTITY x "header1.xml">]>`);
    expect((JSON.parse(read(parts)) as Any).page?.header).toBeUndefined();
  });

  it.each(Object.keys(ENCODINGS))("%s: reads no settings from a part declaring a DOCTYPE", (encoding) => {
    const parts = unzipSync(Agent.toBytes({ page: { header: { left: "odd" }, evenHeader: { left: "even" } }, blocks: [] } as Any)) as Record<string, string | Uint8Array>;
    expect((JSON.parse(read(parts)) as Any).page).toHaveProperty("evenHeader");
    parts["word/settings.xml"] = ENCODINGS[encoding]!(withDoctype(text(parts["word/settings.xml"]!)));
    expect((JSON.parse(read(parts)) as Any).page).not.toHaveProperty("evenHeader");
  });

  const dir = mkdtempSync(join(tmpdir(), "lw-dt-"));
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  it("reads nothing of a file a SYSTEM entity names", () => {
    const scratch = join(dir, "scratch.txt");
    writeFileSync(scratch, "SCRATCHFILE");
    const doctype = `<!DOCTYPE r [<!ENTITY x SYSTEM "file:///${scratch.replace(/\\/g, "/")}">]>`;
    for (const part of PARTS) {
      const parts = docx();
      parts[part] = withDoctype(text(parts[part]!).split(MARKER).join("&x;"), doctype);
      expect(read(parts)).not.toContain("SCRATCHFILE");
    }
  });

  it.each(["content.xml", "styles.xml"])("%s: refuses a DOCTYPE in an ODT part, by name", (part) => {
    const odt = unzipSync(new Uint8Array(readFileSync(new URL("../test/fixtures/formats/report.odt", import.meta.url)))) as Record<string, string | Uint8Array>;
    for (const encoding of Object.keys(ENCODINGS)) {
      const copy = { ...odt };
      copy[part] = ENCODINGS[encoding]!(withDoctype(text(odt[part]!)));
      expect(() => Agent.read(zip(copy)), encoding).toThrow(/DOCTYPE/);
    }
  });

  it.each(["word/styles.xml", "word/theme/theme1.xml"])("%s: refuses a DOCTYPE in a template part it carries", (part) => {
    const template = docx();
    template["word/theme/theme1.xml"] ??= `<?xml version="1.0"?><a:theme xmlns:a="urn:a" name="t"/>`;
    for (const encoding of Object.keys(ENCODINGS)) {
      const copy = { ...template };
      copy[part] = ENCODINGS[encoding]!(withDoctype(text(template[part]!)));
      expect(() => Agent.toBytes({ blocks: [] } as Any, { template: zip(copy) } as Any), encoding).toThrow(/DOCTYPE/);
    }
  });
});
