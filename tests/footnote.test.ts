import { describe, expect, it } from "vitest";
import { Agent, DocxWriter, unzipSync, zipSync } from "../src";

/**
 * Footnotes (MOIC, 2026-11-01): writer agents write documents WITH footnotes,
 * beside the headers, footers and page numbers they already write. Mirrors PHP
 * last-word 0.15.0 `tests/Unit/FootnoteTest.php`, case for case.
 *
 * A footnote is a RUN: `{"text": "", "footnote": {"blocks": [...]}}`, placed
 * where the reference mark goes. Word numbers the marks; the model carries no
 * number, so an agent never has to keep a counter in step with the text. The
 * writer emits word/footnotes.xml with the separator and continuation
 * separator Word requires, a FootnoteReference-styled `w:footnoteReference`
 * run, the relationship, the content type and settings `w:footnotePr`; the
 * reader gives the same run back. Without a footnote nothing changes, byte for
 * byte.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const RED_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR42mP4z8AARAwQCgAf7gP9Y167WwAAAABJRU5ErkJggg==";

const dec = new TextDecoder();
const enc = new TextEncoder();

function fnNote(text: string): Any {
  return { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text }] }] } };
}

function fnDoc(): Any {
  return {
    title: "Footnotes",
    blocks: [
      { type: "paragraph", runs: [{ text: "Revenue grew 18%" }, fnNote("Audited figures, FY2026."), { text: " on the prior year." }] },
    ],
  };
}

function partsOf(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = dec.decode(data);
  return out;
}

function fnParts(doc: Any): Record<string, string> {
  return partsOf(Agent.toBytes(doc));
}

function fnPart(bytes: Uint8Array, name: string): string {
  return partsOf(bytes)[name]!;
}

function fnReplacePart(bytes: Uint8Array, name: string, xml: string): Uint8Array {
  const files = unzipSync(bytes);
  files[name] = enc.encode(xml);
  return zipSync(Object.entries(files).map(([n, data]) => ({ name: n, data })));
}

/**
 * The package is self-consistent: every internal relationship target exists, and
 * every part has a content type. The JS counterpart of PHP's
 * `Package::assertConsistent()`.
 */
function assertConsistent(parts: Record<string, string>): void {
  const types = parts["[Content_Types].xml"]!;
  for (const name of Object.keys(parts)) {
    if (name === "[Content_Types].xml") continue;
    const ext = name.split(".").pop()!;
    const declared = types.includes(`PartName="/${name}"`) || types.includes(`Extension="${ext}"`);
    expect(declared, `${name} has no content type`).toBe(true);
  }
  for (const [name, xml] of Object.entries(parts)) {
    if (!name.endsWith(".rels")) continue;
    const base = name.startsWith("_rels/") ? "" : name.slice(0, name.indexOf("_rels/"));
    for (const m of xml.matchAll(/<Relationship [^>]*?Target="([^"]+)"([^>]*)\/>/g)) {
      if (m[0].includes('TargetMode="External"')) continue;
      expect(parts, `${name} points at a missing ${base}${m[1]}`).toHaveProperty([`${base}${m[1]}`]);
    }
  }
}

// ─── Schema + validation ─────────────────────────────────────────────────

describe("footnotes: schema and validation", () => {
  it("publishes a footnote on the run, holding blocks", () => {
    const schema = Agent.jsonSchema() as Any;

    expect(schema.definitions.run.properties.footnote).toEqual({ $ref: "#/definitions/footnote" });
    expect(schema.definitions.footnote.required).toEqual(["blocks"]);
    expect(schema.definitions.footnote.properties.blocks.type).toBe("array");
  });

  it("accepts a document with footnotes, including rich ones and ones in tables and lists", () => {
    const doc = fnDoc();
    doc.blocks.push({ type: "list", items: [{ runs: [{ text: "Item" }, fnNote("On a list item.")] }] });
    doc.blocks.push({ type: "table", rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "12.5" }, fnNote("In a cell.")] }] }] }] });
    doc.blocks.push({
      type: "paragraph",
      runs: [
        { text: "Rich" },
        {
          text: "",
          footnote: {
            blocks: [
              { type: "paragraph", runs: [{ text: "Sources:", bold: true }] },
              { type: "list", items: [{ runs: [{ text: "Ledger" }] }, { runs: [{ text: "Board pack" }] }] },
            ],
          },
        },
      ],
    });

    expect(Agent.validate(doc)).toEqual([]);
  });

  const refused: [string, Any, string][] = [
    ["no blocks", { blocks: [{ type: "paragraph", runs: [{ text: "", footnote: {} }] }] }, "/blocks/0/runs/0/footnote/blocks"],
    ["not an object", { blocks: [{ type: "paragraph", runs: [{ text: "", footnote: "Source: X" }] }] }, "/blocks/0/runs/0/footnote"],
    ["text on the mark", { blocks: [{ type: "paragraph", runs: [{ ...fnNote("n"), text: "x" }] }] }, "/blocks/0/runs/0/text"],
    [
      "a footnote in a footnote",
      { blocks: [{ type: "paragraph", runs: [{ text: "", footnote: { blocks: [{ type: "paragraph", runs: [fnNote("deeper")] }] } }] }] },
      "/blocks/0/runs/0/footnote/blocks/0/runs/0/footnote",
    ],
    [
      "a page break in a footnote",
      { blocks: [{ type: "paragraph", runs: [{ text: "", footnote: { blocks: [{ type: "pageBreak" }] } }] }] },
      "/blocks/0/runs/0/footnote/blocks/0/type",
    ],
    [
      "a footnote in a footer",
      { page: { footer: { blocks: [{ type: "paragraph", runs: [fnNote("n")] }] } }, blocks: [] },
      "/page/footer/blocks/0/runs/0/footnote",
    ],
  ];
  for (const [name, doc, path] of refused) {
    it(`refuses a footnote that is malformed, misplaced or nested: ${name}`, () => {
      expect(Agent.validate(doc).map((e) => e.path)).toContain(path);
    });
  }

  it("repairs the near misses an agent makes, and says so", () => {
    const doc = {
      blocks: [
        {
          type: "paragraph",
          runs: [
            { footnote: { blocks: [{ type: "paragraph", runs: [{ text: "No text key." }] }] } },
            { text: "Typed on the mark", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Kept." }] }] } },
          ],
        },
      ],
    };

    const out = Agent.validateAndRepair(doc);

    expect(out.ok).toBe(true);
    expect((out.schema.blocks[0] as Any).runs).toStrictEqual([
      { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "No text key." }] }] } },
      { text: "Typed on the mark" },
      { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Kept." }] }] } },
    ]);
    expect(out.errors).not.toEqual([]);
  });

  it("repairs away a footnote where Word cannot hold one", () => {
    const doc = {
      page: { header: { blocks: [{ type: "paragraph", runs: [{ text: "Head" }, fnNote("n")] }] } },
      blocks: [
        {
          type: "paragraph",
          runs: [{ text: "Body" }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Outer" }, fnNote("inner")] }] } }],
        },
      ],
    };

    const out = Agent.validateAndRepair(doc) as Any;

    expect(out.ok).toBe(true);
    expect(out.schema.page.header.blocks[0].runs).toStrictEqual([{ text: "Head" }]);
    expect(out.schema.blocks[0].runs[1].footnote.blocks[0].runs).toStrictEqual([{ text: "Outer" }]);
  });

  it("wraps a string or a bare list of blocks into a footnote", () => {
    const out = Agent.validateAndRepair({
      blocks: [{ type: "paragraph", runs: [{ text: "", footnote: "Source: X" }, { text: "", footnote: [{ type: "paragraph", runs: [{ text: "Y" }] }] }] }],
    }) as Any;

    expect(out.ok).toBe(true);
    expect(out.schema.blocks[0].runs).toStrictEqual([fnNote("Source: X"), fnNote("Y")]);
  });
});

// ─── Writer ──────────────────────────────────────────────────────────────

describe("footnotes: writer", () => {
  it("writes footnotes.xml with the separators Word requires, wired in everywhere Word looks", () => {
    const parts = fnParts(fnDoc());

    expect(parts).toHaveProperty(["word/footnotes.xml"]);
    const fn = parts["word/footnotes.xml"]!;

    expect(fn).toContain('<w:footnote w:type="separator" w:id="-1">');
    expect(fn).toContain("<w:separator/>");
    expect(fn).toContain('<w:footnote w:type="continuationSeparator" w:id="0">');
    expect(fn).toContain("<w:continuationSeparator/>");
    expect(fn).toContain('<w:footnote w:id="1">');
    expect(fn).toContain('<w:pStyle w:val="FootnoteText"/>');
    expect(fn).toContain('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>');
    expect(fn).toContain("Audited figures, FY2026.");
    expect(parts["word/document.xml"]).toContain('<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>');
    expect(parts["word/_rels/document.xml.rels"]).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"',
    );
    expect(parts["[Content_Types].xml"]).toContain(
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
    );
    expect(parts["word/settings.xml"]).toContain('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>');
    expect(parts["word/styles.xml"]).toContain('w:styleId="FootnoteText"');
    expect(parts["word/styles.xml"]).toContain('w:styleId="FootnoteReference"');
    expect(parts["word/styles.xml"]).toContain('<w:vertAlign w:val="superscript"/>');

    assertConsistent(parts);
  });

  it("numbers footnotes in document order, wherever they sit", () => {
    const doc = fnDoc();
    doc.blocks.push({ type: "table", rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "c" }, fnNote("Second.")] }] }] }] });
    doc.blocks.push({ type: "heading", level: 2, runs: [{ text: "H" }, fnNote("Third.")] });

    const parts = fnParts(doc);
    const refs = [...parts["word/document.xml"]!.matchAll(/<w:footnoteReference w:id="(\d+)"\/>/g)].map((m) => m[1]);
    const notes = [...parts["word/footnotes.xml"]!.matchAll(/<w:footnote w:id="(\d+)">[\s\S]*?<\/w:footnote>/g)];

    expect(refs).toEqual(["1", "2", "3"]);
    expect(notes.map((m) => m[1])).toEqual(["1", "2", "3"]);
    expect(notes[1]![0]).toContain("Second.");
    expect(notes[2]![0]).toContain("Third.");
  });

  it("gives a footnote its own relationships for its links and images", () => {
    const doc = {
      blocks: [
        {
          type: "paragraph",
          runs: [
            { text: "See" },
            {
              text: "",
              footnote: {
                blocks: [
                  { type: "paragraph", runs: [{ text: "the filing", link: "https://example.com/10-k" }] },
                  { type: "image", src: RED_PNG, alt: "seal" },
                ],
              },
            },
            { text: " and " },
            { text: "this link", link: "https://example.org/" },
          ],
        },
      ],
    };

    const parts = fnParts(doc);

    expect(parts).toHaveProperty(["word/_rels/footnotes.xml.rels"]);
    expect(parts["word/_rels/footnotes.xml.rels"]).toContain('Target="https://example.com/10-k"');
    expect(parts["word/_rels/footnotes.xml.rels"]).toContain('Target="media/image1.png"');
    expect(parts["word/_rels/document.xml.rels"]).not.toContain("example.com/10-k");
    expect(parts["word/_rels/document.xml.rels"]).toContain('Target="https://example.org/"');

    assertConsistent(parts);
  });

  it("puts the mark first even when the footnote opens with a table", () => {
    const doc = {
      blocks: [
        {
          type: "paragraph",
          runs: [
            { text: "x" },
            { text: "", footnote: { blocks: [{ type: "table", rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "cell" }] }] }] }] }] } },
          ],
        },
      ],
    };

    const fn = fnParts(doc)["word/footnotes.xml"]!;
    const m = /<w:footnote w:id="1">([\s\S]*?)<\/w:footnote>/.exec(fn)!;

    // The mark paragraph comes before the table; a footnote must also END in a paragraph.
    expect(
      m[1]!.startsWith(
        '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>',
      ),
    ).toBe(true);
    expect(m[1]!.endsWith("</w:p>")).toBe(true);
  });

  it("changes nothing about a document without footnotes", () => {
    const parts = fnParts({ blocks: [{ type: "paragraph", runs: [{ text: "Plain" }] }] });

    expect(parts).not.toHaveProperty(["word/footnotes.xml"]);
    expect(parts).not.toHaveProperty(["word/settings.xml"]);
    expect(parts["word/styles.xml"]).not.toContain("Footnote");
    expect(parts["[Content_Types].xml"]).not.toContain("footnotes");
    expect(parts["word/_rels/document.xml.rels"]).not.toContain("footnotes");
  });

  it("writes nothing for a footnote in a header, which Word cannot hold", () => {
    // Validation refuses this document, so the writer is driven directly: what
    // it does with unvalidated input is the claim.
    const doc = { page: { header: { blocks: [{ type: "paragraph", runs: [{ text: "Head" }, fnNote("n")] }] } }, blocks: [] };

    const parts = partsOf(new DocxWriter().toBytes(doc as Any));

    expect(parts).not.toHaveProperty(["word/footnotes.xml"]);
    for (const [name, xml] of Object.entries(parts)) {
      if (name.startsWith("word/header")) expect(xml).not.toContain("footnote");
    }
  });

  it("writes nothing for a footnote inside a footnote, and still writes the outer one", () => {
    const doc = {
      blocks: [{ type: "paragraph", runs: [{ text: "Body" }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Outer" }, fnNote("inner")] }] } }] }],
    };

    const parts = partsOf(new DocxWriter().toBytes(doc as Any));

    expect([...parts["word/footnotes.xml"]!.matchAll(/<w:footnote w:id="\d+">/g)]).toHaveLength(1);
    expect(parts["word/footnotes.xml"]).not.toContain("inner");
    expect(parts["word/footnotes.xml"]).not.toContain("footnoteReference");
  });

  it("writes text typed on the mark before it rather than losing it", () => {
    const doc = { blocks: [{ type: "paragraph", runs: [{ ...fnNote("n"), text: "Typed" }] }] };

    const xml = partsOf(new DocxWriter().toBytes(doc as Any))["word/document.xml"]!;

    expect(xml).toContain('<w:r><w:t xml:space="preserve">Typed</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>');
  });
});

// ─── Reader ──────────────────────────────────────────────────────────────

describe("footnotes: reader", () => {
  it("reads its own footnotes back exactly, wherever they sit", () => {
    const doc = {
      title: "Footnotes",
      blocks: [
        { type: "heading", level: 1, runs: [{ text: "Results" }, fnNote("Unaudited.")] },
        { type: "paragraph", runs: [{ text: "Revenue grew " }, { text: "18%", bold: true }, fnNote("Against FY2025."), { text: " on the prior year." }] },
        { type: "list", items: [{ runs: [{ text: "EMEA" }, fnNote("Includes Turkey.")] }] },
        {
          type: "table",
          rows: [
            {
              cells: [
                { blocks: [{ type: "paragraph", runs: [{ text: "Q3" }, fnNote("Thirteen weeks.")] }] },
                { blocks: [{ type: "paragraph", runs: [{ text: "1,250" }] }] },
              ],
            },
          ],
        },
        { type: "quote", blocks: [{ type: "paragraph", runs: [{ text: "Quoted" }, fnNote("Interview, 3 October.")] }] },
        { type: "paragraph", runs: [{ text: "Two in a row" }, fnNote("First."), fnNote("Second.")] },
        {
          type: "paragraph",
          runs: [
            { text: "Rich" },
            {
              text: "",
              footnote: {
                blocks: [
                  { type: "paragraph", runs: [{ text: "Sources: ", bold: true }, { text: "the ledger", link: "https://example.com/ledger" }] },
                  { type: "list", ordered: true, items: [{ runs: [{ text: "Ledger" }] }, { runs: [{ text: "Board pack" }] }] },
                  { type: "paragraph", runs: [{ text: " leading space kept" }] },
                ],
              },
            },
          ],
        },
      ],
    };

    expect(Agent.read(Agent.toBytes(doc))).toStrictEqual(doc);
  });

  it("keeps two identical footnotes in a row as two", () => {
    // Two different footnotes never merge -- their `footnote` values differ --
    // so only identical ones test that a reader never merges footnote runs.
    const doc = { blocks: [{ type: "paragraph", runs: [{ text: "Twice" }, fnNote("Ibid."), fnNote("Ibid.")] }] };

    expect(Agent.read(Agent.toBytes(doc))).toStrictEqual(doc);
  });

  it("reads a footnote that is the whole paragraph", () => {
    const doc = { blocks: [{ type: "paragraph", runs: [fnNote("Only a mark.")] }] };

    expect(Agent.read(Agent.toBytes(doc))).toStrictEqual(doc);
  });

  it("reads Word's footnotes: any ids, the mark and its space, the separators skipped", () => {
    let bytes = Agent.toBytes({
      blocks: [{ type: "paragraph", runs: [{ text: "Body" }, fnNote("placeholder"), { text: " more" }, fnNote("placeholder 2")] }],
    });

    // Word 2003-era numbering: separators at 0 and 1, notes from 2, and a
    // reference written out of order with its definition.
    const word =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="7"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Seventh, </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Second.</w:t></w:r></w:p></w:footnote>' +
      "</w:footnotes>";
    bytes = fnReplacePart(bytes, "word/footnotes.xml", word);
    // One pass, so 1 -> 2 is not then rewritten 2 -> 7.
    const swap: Record<string, string> = {
      '<w:footnoteReference w:id="1"/>': '<w:footnoteReference w:id="2"/>',
      '<w:footnoteReference w:id="2"/>': '<w:footnoteReference w:id="7"/>',
    };
    bytes = fnReplacePart(
      bytes,
      "word/document.xml",
      fnPart(bytes, "word/document.xml").replace(/<w:footnoteReference w:id="[12]"\/>/g, (s) => swap[s]!),
    );

    expect(Agent.read(bytes).blocks).toStrictEqual([
      {
        type: "paragraph",
        runs: [
          { text: "Body" },
          fnNote("Second."),
          { text: " more" },
          { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Seventh, " }, { text: "italic", italic: true }] }] } },
        ],
      },
    ]);
  });

  it("reads a reference whose footnote is missing as nothing, and never invents one", () => {
    let bytes = Agent.toBytes(fnDoc());
    bytes = fnReplacePart(bytes, "word/document.xml", fnPart(bytes, "word/document.xml").replace('w:id="1"', 'w:id="99"'));

    expect((Agent.read(bytes).blocks[0] as Any).runs).toStrictEqual([{ text: "Revenue grew 18% on the prior year." }]);
  });

  it("reads a run that holds text and a reference as two runs", () => {
    const bytes = Agent.toBytes(fnDoc());
    // Another producer's spelling: text and the reference in ONE w:r.
    const xml = fnPart(bytes, "word/document.xml").replace(
      '<w:r><w:t xml:space="preserve">Revenue grew 18%</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>',
      '<w:r><w:t xml:space="preserve">Revenue grew 18%</w:t><w:footnoteReference w:id="1"/></w:r>',
    );
    expect(xml).toContain("18%</w:t><w:footnoteReference");

    expect(Agent.read(fnReplacePart(bytes, "word/document.xml", xml)).blocks).toStrictEqual(fnDoc().blocks);
  });

  it("reads a footnote that opens or ends with a table without the writer's mark and closing paragraphs", () => {
    // The writer gives such a note a mark-only paragraph before the table and an
    // empty one after it; neither is the author's, and PHP reads neither.
    const table = { type: "table", rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "cell" }] }] }] }] };
    const doc = {
      blocks: [
        { type: "paragraph", runs: [{ text: "Opens" }, { text: "", footnote: { blocks: [table] } }] },
        { type: "paragraph", runs: [{ text: "Ends" }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "Before" }] }, table] } }] },
      ],
    };

    expect(Agent.read(Agent.toBytes(doc))).toStrictEqual(doc);
  });

  it("diffs a document against its own read-back as no change", () => {
    const doc = fnDoc();

    expect(Agent.diff(doc, Agent.read(Agent.toBytes(doc)))).toEqual([]);
  });

  it("does not read a footnotes part that declares a DOCTYPE", () => {
    let bytes = Agent.toBytes(fnDoc());
    bytes = fnReplacePart(bytes, "word/footnotes.xml", '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]>' + fnPart(bytes, "word/footnotes.xml").replace(/^<\?xml[^>]*\?>/, ""));

    expect((Agent.read(bytes).blocks[0] as Any).runs).toStrictEqual([{ text: "Revenue grew 18% on the prior year." }]);
  });
});

// ─── Exports (not covered yet, as in PHP 0.15.0) ─────────────────────────

describe("footnotes: exports", () => {
  it("markdown does not carry a footnote: the mark vanishes and its text is not exported", () => {
    const md = Agent.toMarkdown(fnDoc());

    expect(md).toContain("Revenue grew 18% on the prior year.");
    expect(md).not.toContain("Audited");
  });
});
