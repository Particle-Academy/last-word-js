import { readFileSync } from "node:fs";
import { deflateRawSync } from "node:zlib";
import { describe, expect, it } from "vitest";
import { Agent, UnsupportedFormatException, zipSync } from "../src";
import * as L from "./support/legacy-files";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/**
 * `.doc`, `.odt` and `.rtf` read as the SAME document a `.docx` does
 * (last-word#1). Mirrors PHP `LegacyFormatsTest`.
 *
 * `test/fixtures/formats/` holds byte-identical copies of the PHP package's
 * fixtures: one document written by last-word and converted by LibreOffice.
 * Each read is held to `report.read.json`, which the PHP and Python engines
 * assert too, as JSON text, so key order and every value must agree across the
 * three runtimes, not only within this one.
 */

const fixture = (ext: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`../test/fixtures/formats/report.${ext}`, import.meta.url)));

const expectedText = (): string =>
  readFileSync(new URL("../test/fixtures/formats/report.read.json", import.meta.url), "utf8");

/** The read as PHP's JSON_PRETTY_PRINT writes it, which is how report.read.json was made. */
const asJson = (doc: unknown): string => JSON.stringify(doc, null, 4) + "\n";

function texts(blocks: Any[]): string[] {
  const out: string[] = [];
  const items = (list: Any[]): void => {
    for (const item of list) {
      out.push(...item.runs.map((r: Any) => r.text));
      items(item.children ?? []);
    }
  };
  for (const block of blocks) {
    if (block.runs) out.push(...block.runs.map((r: Any) => r.text));
    else if (block.type === "list") items(block.items);
    else if (block.type === "table") {
      for (const row of block.rows) for (const cell of row.cells) out.push(...texts(cell.blocks));
    }
  }
  return out;
}

function thrown(read: () => unknown): Any {
  try {
    read();
  } catch (e) {
    return e;
  }
  throw new Error("expected the read to throw, and it returned");
}

describe("one document, four formats, one answer", () => {
  it("reads the docx exactly as report.read.json, the answer all three runtimes assert", () => {
    expect(asJson(Agent.read(fixture("docx")))).toBe(expectedText());
  });

  it("reads the docx as the source document, so agreeing with it means something", () => {
    const doc: Any = Agent.read(fixture("docx"));
    expect(doc.blocks.map((b: Any) => b.type)).toEqual([
      "heading", "paragraph", "heading", "paragraph", "heading", "list", "list", "heading", "table", "paragraph",
    ]);
    expect(doc.blocks[5].items[1].children[0].children[0].runs[0].text).toBe("Third level");
    expect(doc.blocks[8].rows).toHaveLength(3);
  });

  it("reads the legacy .doc as exactly the docx", () => {
    expect(asJson(Agent.read(fixture("doc")))).toBe(expectedText());
  });

  it("reads the .odt as exactly the docx", () => {
    expect(asJson(Agent.read(fixture("odt")))).toBe(expectedText());
  });

  it("reads the .rtf as the docx, less the header-row flag the file does not carry", () => {
    // LibreOffice writes no \trhdr, so there is nothing in this file to recover.
    expect(new TextDecoder().decode(fixture("rtf"))).not.toContain(String.raw`\trhdr`);

    const expected: Any = JSON.parse(expectedText());
    delete expected.blocks[8].rows[0].header;

    expect(asJson(Agent.read(fixture("rtf")))).toBe(asJson(expected));
  });

  it("recovers the text that is hardest to get right in every format", () => {
    for (const format of ["doc", "odt", "rtf"]) {
      const text = texts((Agent.read(fixture(format)) as Any).blocks).join("");
      expect(text).toContain("Café, naïve, jalapeño — 日本語のテキスト and an emoji 🎉 in one line.");
      expect(text).toContain("São Paulo");
      expect(text).toContain("−4.2%");
      expect(text).not.toContain("HYPERLINK");
      expect(text).not.toContain("Hyperlink");
      expect(text).not.toContain("Times New Roman");
    }
  });
});

describe("a format it still cannot read is refused by name", () => {
  it("names an .xls, which is a compound file but not a Word document", () => {
    const e = thrown(() => Agent.read(L.cfb({ Workbook: L.ascii("cells") })));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("xls");
    expect(e.message).toContain("Excel");
  });

  it("names a compound file it does not recognise at all", () => {
    const e = thrown(() => Agent.read(L.cfb({ Contents: L.ascii("something") })));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("cfb");
  });

  it("names a Word 95 file, which predates the binary format it reads", () => {
    const e = thrown(() => Agent.read(L.cfb(L.word(L.ascii("Old\r"), { nFib: 0x0065 }))));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("doc");
    expect(e.message).toContain("Word 95");
  });

  it("names an encrypted .doc rather than reading ciphertext as text", () => {
    const e = thrown(() => Agent.read(L.cfb(L.word(L.ascii("Secret\r"), { flags: 0x0100 }))));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("doc");
    expect(e.message).toContain("password");
  });

  it.each([
    ["xlsx", { "[Content_Types].xml": "<Types/>", "xl/workbook.xml": "<workbook/>" }],
    ["pptx", { "[Content_Types].xml": "<Types/>", "ppt/presentation.xml": "<presentation/>" }],
    ["ods", { mimetype: "application/vnd.oasis.opendocument.spreadsheet", "content.xml": "<x/>" }],
  ])("names an Office zip that is not a word-processing document: %s", (format, entries) => {
    const e = thrown(() => Agent.read(L.zip(entries)));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe(format);
  });

  it("refuses bytes that are no document with the same exception type", () => {
    const e = thrown(() => Agent.read(L.ascii("just some text\n")));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("unknown");
  });
});

describe("a hand-built .doc", () => {
  it("reads its paragraphs, which is what the guard tests below break", () => {
    expect((Agent.read(L.cfb(L.word(L.ascii("Hello\rWorld\r")))) as Any).blocks).toEqual([
      { type: "paragraph", runs: [{ text: "Hello" }] },
      { type: "paragraph", runs: [{ text: "World" }] },
    ]);
  });

  it("keeps a field result and drops its instruction", () => {
    const text = L.ascii("See \x13 PAGE \x147\x15 now\r");
    expect((Agent.read(L.cfb(L.word(text))) as Any).blocks[0].runs).toEqual([{ text: "See 7 now" }]);
  });

  it("decodes an 8-bit piece as Windows-1252, not Latin-1", () => {
    const text = Uint8Array.from([0x93, 0x48, 0x69, 0x94, 0x20, 0x80, 0x0d]);
    expect((Agent.read(L.cfb(L.word(text))) as Any).blocks[0].runs[0].text).toBe("“Hi” €");
  });

  it("joins a UTF-16 surrogate pair and replaces half of one", () => {
    const units = [0xd83c, 0xdf89, 0x20, 0xd800, 0x41, 0x0d];
    const text = Uint8Array.from(units.flatMap((u) => [u & 0xff, u >> 8]));
    expect((Agent.read(L.cfb(L.word(text, { unicode: true }))) as Any).blocks[0].runs[0].text).toBe(
      "🎉 " + String.fromCharCode(0xfffd) + "A",
    );
  });
});

describe("damaged and hostile files fail fast, and say what is wrong", () => {
  // A damaged file is not an unsupported format: a person needs to be told
  // "this file is broken", not "save it as .docx".
  const damaged = (bytes: Uint8Array, message: string): void => {
    const e = thrown(() => Agent.read(bytes));
    expect(e).toBeInstanceOf(Error);
    expect(e).not.toBeInstanceOf(UnsupportedFormatException);
    expect(e.message).toContain(message);
  };
  const hello = (): Uint8Array => L.cfb(L.word(L.ascii("Hello\r")));

  it("refuses a compound file cut off inside its header", () => {
    const bytes = new Uint8Array(208);
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    damaged(bytes, "header is missing or truncated");
  });

  it("refuses a sector chain that loops", () => {
    damaged(L.patch(hello(), L.sectorOffset(0) + 4, L.le32(1)), "directory chain loops");
  });

  it("refuses a chain that points outside the file", () => {
    damaged(L.patch(hello(), 48, L.le32(100000)), "points outside the file");
  });

  it("refuses a stream whose declared size its chain cannot hold", () => {
    damaged(L.patch(hello(), L.entryOffset(1) + 120, L.le32(64 * 1024 * 1024)), "shorter than its declared size");
  });

  it("refuses a stream size too large to read, before reading it", () => {
    damaged(L.patch(hello(), L.entryOffset(1) + 120, L.le32(0xf0000000)), "too large to read");
  });

  it("refuses a DIFAT chain that loops", () => {
    let bytes = L.patch(hello(), 68, L.le32(2));
    bytes = L.patch(bytes, L.sectorOffset(2), L.le32(...new Array(127).fill(L.FREESECT), 2));
    damaged(bytes, "DIFAT chain loops");
  });

  it("refuses an allocation table that lists more sectors than the file holds", () => {
    damaged(L.patch(hello(), 76, L.le32(...new Array(109).fill(0))), "more sectors than the file holds");
  });

  it("stops walking a directory whose siblings form a cycle", () => {
    let bytes = L.cfb({ Contents: L.ascii("a"), Other: L.ascii("b") });
    bytes = L.patch(bytes, L.entryOffset(1) + 68, L.le32(2));
    bytes = L.patch(bytes, L.entryOffset(2) + 72, L.le32(1));

    const e = thrown(() => Agent.read(bytes));
    expect(e).toBeInstanceOf(UnsupportedFormatException);
    expect(e.format).toBe("cfb");
  });

  it("refuses a piece table whose property records do not move forward", () => {
    const prc = Uint8Array.from([0x01, 0xfd, 0xff]);
    const clx = L.piecesClx([[0, 6, 1024, true]], prc);
    damaged(L.cfb(L.word(L.ascii("Hello\r"), { clx })), "piece table is malformed");
  });

  it("reads overlapping pieces once, not once per piece", () => {
    const clx = L.piecesClx([[0, 6, 1024, true], [6, 0, 1024, true], [0, 6, 1024, true]]);
    const doc: Any = Agent.read(L.cfb(L.word(L.ascii("Hello\r"), { clx, ccpText: 12 })));
    expect(doc.blocks).toEqual([{ type: "paragraph", runs: [{ text: "Hello" }] }]);
  });

  it("bounds a text length of four billion characters by the bytes present", () => {
    const clx = L.piecesClx([[0, 0xfffffff0, 1024, true]]);
    const doc: Any = Agent.read(L.cfb(L.word(L.ascii("Hello\r"), { clx, ccpText: 0xfffffff0 })));
    expect(texts(doc.blocks)[0]).toBe("Hello");
  });

  it("refuses an ODT part carrying a DOCTYPE, the door to entity expansion", () => {
    const xml =
      '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>';
    damaged(L.odt("", xml), "DOCTYPE");
  });

  it("parses an ODT part nested 257 elements deep and refuses one nested 258", () => {
    // libxml's limit without XML_PARSE_HUGE, which the PHP engine inherits. The
    // part's own wrappers are four deep (document-content, body, text, p).
    const nested = (spans: number): Uint8Array =>
      L.odt("<text:p>" + "<text:span>".repeat(spans) + "deep" + "</text:span>".repeat(spans) + "</text:p>");

    expect(texts((Agent.read(nested(253)) as Any).blocks)).toEqual(["deep"]);
    damaged(nested(254), "Could not parse content.xml");
  });

  it("caps an ODT space run of two billion", () => {
    const doc: Any = Agent.read(L.odt('<text:p>a<text:s text:c="2000000000"/>b</text:p>'));
    expect(doc.blocks[0].runs[0].text).toHaveLength(1002);
  });

  it("caps the cells ODT repeat attributes can add in total", () => {
    const row =
      '<table:table-row table:number-rows-repeated="1000">' +
      '<table:table-cell table:number-columns-repeated="1000"><text:p>x</text:p></table:table-cell>' +
      "</table:table-row>";
    const doc: Any = Agent.read(L.odt("<table:table>" + row.repeat(5) + "</table:table>"));
    const cells = doc.blocks[0].rows.reduce((n: number, r: Any) => n + r.cells.length, 0);
    expect(cells).toBeLessThanOrEqual(5 * 1000 + 100_000 + 1000);
  });

  it("stops inflating an ODT part at 64 MB whatever its entry declares", () => {
    // A zip bomb: 65 MB of spaces deflate to a few kilobytes, and the entry
    // claims to be tiny. Only an inflater that counts its output stops it.
    const huge = deflateRawSync(Buffer.alloc(65 * 1024 * 1024, 0x20));
    const bytes = bombZip("content.xml", new Uint8Array(huge), 100);
    damaged(bytes, "ODT part content.xml is too large to read");
  });

  it("refuses a truncated DEFLATE stream rather than inventing the missing bytes", () => {
    const deflated = new Uint8Array(deflateRawSync(Buffer.from("<office:document-content/>".repeat(50))));
    damaged(bombZip("content.xml", deflated.subarray(0, 4), 1300), "Could not open the ODT archive");
  });

  it("refuses RTF nested past any real document", () => {
    damaged(L.ascii(String.raw`{\rtf1 ` + "{".repeat(20000) + "x" + "}".repeat(20000) + "}"), "nests groups too deeply");
  });

  it("skips \\bin data by its length without reading past the end", () => {
    const doc: Any = Agent.read(L.ascii(String.raw`{\rtf1 before\par{\bin4000000000 abc}}`));
    expect(texts(doc.blocks)).toEqual(["before"]);
  });

  it("survives unbalanced RTF braces", () => {
    expect((Agent.read(L.ascii(String.raw`{\rtf1 {\b one}\par}}}two\par`)) as Any).blocks).toEqual([
      { type: "paragraph", runs: [{ text: "one", bold: true }] },
      { type: "paragraph", runs: [{ text: "two" }] },
    ]);
  });
});

/** An ODT-looking zip whose one deflated entry declares `declared` bytes. */
function bombZip(name: string, deflated: Uint8Array, declared: number): Uint8Array {
  const stored = zipSync([
    { name: "mimetype", data: new TextEncoder().encode("application/vnd.oasis.opendocument.text") },
    { name, data: deflated },
  ]);
  // zipSync STOREs; mark the second entry DEFLATE-compressed and set its sizes.
  const out = stored.slice();
  const dv = new DataView(out.buffer);
  const secondLocal = 30 + "mimetype".length + "application/vnd.oasis.opendocument.text".length;
  dv.setUint16(secondLocal + 8, 8, true);
  dv.setUint32(secondLocal + 22, declared, true);
  const eocd = out.length - 22;
  let cd = dv.getUint32(eocd + 16, true);
  cd += 46 + dv.getUint16(cd + 28, true) + dv.getUint16(cd + 30, true) + dv.getUint16(cd + 32, true);
  dv.setUint16(cd + 10, 8, true);
  dv.setUint32(cd + 24, declared, true);
  return out;
}
