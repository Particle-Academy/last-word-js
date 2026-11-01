import { describe, expect, it } from "vitest";
import { Agent, Schema, unzipSync, zipSync } from "../src";

/**
 * First and even bands of a document's own (PHP last-word 0.16.0, MOIC
 * 2026-11-01). Mirrors PHP's FirstPageBandsTest, less what this engine does not
 * have: it writes no `{blocks}` band and carries no template section
 * (`templateParts`), so the letterhead option `templateFirstPageOnly` has
 * nothing to act on here. Byte parity with PHP is in parity.test.ts.
 *
 * - `page.firstHeader` / `page.firstFooter` / `page.evenHeader` /
 *   `page.evenFooter` are variants, each the shape of `page.header`.
 *   `page.header` stays "every page": with a `firstHeader` it is every page but
 *   the first, with an `evenHeader` the odd pages.
 * - Parts are numbered by SLOT: `{kind}1.xml` default, `2` first, `3` even.
 * - read() returns the variants while the switch that makes Word show them is on.
 */
type Any = Record<string, unknown>;

function fpDoc(page: Any = {}): Any {
  const doc: Any = { blocks: [{ type: "paragraph", runs: [{ text: "A letter of several pages." }] }] };
  if (Object.keys(page).length > 0) doc.page = page;
  return doc;
}

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = new TextDecoder().decode(data);
  return out;
}

function rezip(p: Record<string, string>): Uint8Array {
  return zipSync(Object.entries(p).map(([name, xml]) => ({ name, data: new TextEncoder().encode(xml) })));
}

const write = (doc: Any): Record<string, string> => parts(Agent.toBytes(doc as never));

/** The section's header/footer references, in order, as "kind:type:rId". */
function refs(p: Record<string, string>): string[] {
  return [...p["word/document.xml"]!.matchAll(/<w:(header|footer)Reference w:type="([a-z]+)" r:id="([^"]+)"\/>/g)].map(
    (m) => `${m[1]}:${m[2]}:${m[3]}`,
  );
}

/** Relationship id => target, for the header and footer relationships. */
function targets(p: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of p["word/_rels/document.xml.rels"]!.matchAll(/Id="([^"]+)" Type="[^"]+\/(?:header|footer)" Target="([^"]+)"/g)) {
    out[m[1]!] = m[2]!;
  }
  return out;
}

/** Every part the package declares and references is present, and vice versa. */
function assertConsistent(p: Record<string, string>): void {
  const declared = [...p["[Content_Types].xml"]!.matchAll(/PartName="\/([^"]+)"/g)].map((m) => m[1]!);
  for (const name of declared) expect(p, `declared ${name} is missing`).toHaveProperty([name]);
  for (const name of Object.keys(p).filter((n) => /^word\/(header|footer)\d+\.xml$/.test(n))) {
    expect(declared, `${name} is not declared`).toContain(name);
    expect(Object.values(targets(p)), `${name} has no relationship`).toContain(name.slice("word/".length));
  }
  for (const id of refs(p).map((r) => r.split(":")[2]!)) expect(targets(p), `${id} has no relationship`).toHaveProperty([id]);
}

describe("a document's own first and even bands", () => {
  it("writes a first-page band as Word's first header, with w:titlePg", () => {
    const p = write(fpDoc({ header: { left: "Every later page" }, firstHeader: { left: "The first page" }, footer: { center: "{page}" } }));

    expect(refs(p)).toEqual([
      "header:default:rIdHeader1",
      "header:first:rIdHeader2",
      // No firstFooter: the footer is on the first page too.
      "footer:default:rIdFooter1",
      "footer:first:rIdFooter1",
    ]);
    expect(targets(p)).toEqual({ rIdHeader1: "header1.xml", rIdHeader2: "header2.xml", rIdFooter1: "footer1.xml" });
    expect(p["word/header2.xml"]).toContain("The first page");
    expect(p["word/document.xml"]).toContain("<w:titlePg/></w:sectPr>");
    expect(p).not.toHaveProperty(["word/settings.xml"]);
    assertConsistent(p);
  });

  it("writes an even-page band with the even/odd setting", () => {
    const p = write(fpDoc({ header: { right: "Odd" }, evenHeader: { left: "Even" }, footer: { center: "{page}" } }));

    expect(refs(p)).toEqual([
      "header:default:rIdHeader1",
      "header:even:rIdHeader3",
      "footer:default:rIdFooter1",
      "footer:even:rIdFooter1",
    ]);
    expect(p["word/header3.xml"]).toContain("Even");
    expect(p["word/settings.xml"]).toContain("<w:evenAndOddHeaders/>");
    expect(p["[Content_Types].xml"]).toContain('PartName="/word/settings.xml"');
    expect(p["word/_rels/document.xml.rels"]).toContain('Id="rIdSettings"');
    expect(p["word/document.xml"]).not.toContain("titlePg");
    assertConsistent(p);
  });

  it("orders the even/odd setting before footnotePr in settings.xml", () => {
    const doc = fpDoc({ header: { right: "Odd" }, evenHeader: { left: "Even" } });
    (doc.blocks as Any[]).push({ type: "paragraph", runs: [{ text: "x" }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "n" }] }] } }] });

    expect(write(doc)["word/settings.xml"]).toContain('<w:evenAndOddHeaders/><w:footnotePr>');
  });

  it("writes a first band with no default as the first page only, and every slot together", () => {
    expect(refs(write(fpDoc({ firstHeader: { left: "Only on page one" } })))).toEqual(["header:first:rIdHeader2"]);

    const p = write(
      fpDoc({
        header: { left: "Odd" },
        footer: { center: "{page}" },
        firstHeader: { left: "First" },
        firstFooter: { center: "Cover" },
        evenHeader: { left: "Even" },
        evenFooter: { center: "Even {page}" },
      }),
    );
    expect(refs(p)).toEqual([
      "header:default:rIdHeader1",
      "header:first:rIdHeader2",
      "header:even:rIdHeader3",
      "footer:default:rIdFooter1",
      "footer:first:rIdFooter2",
      "footer:even:rIdFooter3",
    ]);
    // Parts, content types and relationships in slot order, header before footer.
    expect(Object.keys(p).filter((n) => /^word\/(header|footer)/.test(n))).toEqual([
      "word/header1.xml",
      "word/header2.xml",
      "word/header3.xml",
      "word/footer1.xml",
      "word/footer2.xml",
      "word/footer3.xml",
    ]);
    assertConsistent(p);
  });

  it.each([
    ["slots", { header: { left: "Rest" }, footer: { right: "Page {page}" }, firstHeader: { left: "First", right: "Letterhead" } }],
    ["first only", { firstHeader: { left: "Only on page one" } }],
    ["even", { header: { left: "Odd" }, evenHeader: { left: "Even" }, evenFooter: { center: "Even footer" } }],
    [
      "all three",
      {
        header: { left: "Odd" },
        footer: { center: "{page}" },
        firstHeader: { left: "First" },
        firstFooter: { center: "Cover" },
        evenHeader: { left: "Even" },
        evenFooter: { center: "Even {page}" },
      },
    ],
  ])("round-trips first and even bands: %s", (_label, page) => {
    const read = (Agent.read(Agent.toBytes(fpDoc(page) as never)) as Any).page as Any;

    expect(read).toEqual(page);
    // Keys in schema order: header, footer, firstHeader, firstFooter, evenHeader, evenFooter.
    expect(Object.keys(read)).toEqual(Schema.PAGE_BANDS.filter((k) => k in page));
  });

  it("writes an empty first-page band as a blank first page, and reads it back so", () => {
    const page = { header: { left: "Not on the cover" }, firstHeader: { blocks: [] } };
    const p = write(fpDoc(page));

    expect(refs(p)).toEqual(["header:default:rIdHeader1", "header:first:rIdHeader2"]);
    expect(p["word/header2.xml"]).not.toContain("<w:t");
    expect(p["word/header2.xml"]).toContain("<w:p/>");
    expect((Agent.read(Agent.toBytes(fpDoc(page) as never)) as Any).page).toEqual(page);
    assertConsistent(p);
  });

  it("writes a band with no text as a blank first page too, when the kind has a default", () => {
    const p = write(fpDoc({ footer: { center: "{page}" }, firstFooter: { left: "" } }));

    expect(refs(p)).toEqual(["footer:default:rIdFooter1", "footer:first:rIdFooter2"]);
    expect(p["word/footer2.xml"]).not.toContain("<w:t");
  });

  it("treats an empty first-page band with nothing else as asking for nothing", () => {
    expect(Agent.toBytes(fpDoc({ firstHeader: { blocks: [] }, firstFooter: {} }) as never)).toEqual(
      Agent.toBytes(fpDoc() as never),
    );
  });

  it("writes exactly what 0.12.0 wrote for a document with only header and footer", () => {
    // The pre-0.13.0 shape: one default reference each, no titlePg, no settings.
    const p = write(fpDoc({ header: { left: "H" }, footer: { right: "Page {page} of {pages}" } }));

    expect(refs(p)).toEqual(["header:default:rIdHeader1", "footer:default:rIdFooter1"]);
    expect(p["word/document.xml"]).not.toContain("titlePg");
    expect(p).not.toHaveProperty(["word/settings.xml"]);
  });

  it("ignores a first-page part while w:titlePg is off, as Word does", () => {
    // Our own writer never produces this: a first reference with titlePg off.
    const p = write(fpDoc({ header: { left: "Rest" }, firstHeader: { left: "First" } }));
    p["word/document.xml"] = p["word/document.xml"]!.replace("<w:titlePg/>", "");

    expect((Agent.read(rezip(p)) as Any).page).toEqual({ header: { left: "Rest" } });
  });

  it("reads w:titlePg with w:val off as off", () => {
    const p = write(fpDoc({ header: { left: "Rest" }, firstHeader: { left: "First" } }));
    p["word/document.xml"] = p["word/document.xml"]!.replace("<w:titlePg/>", '<w:titlePg w:val="0"/>');

    expect((Agent.read(rezip(p)) as Any).page).toEqual({ header: { left: "Rest" } });
  });

  it("ignores an even part while settings leave w:evenAndOddHeaders off", () => {
    const p = write(fpDoc({ header: { left: "Odd" }, evenHeader: { left: "Even" } }));
    p["word/settings.xml"] = p["word/settings.xml"]!.replace("<w:evenAndOddHeaders/>", "");

    expect((Agent.read(rezip(p)) as Any).page).toEqual({ header: { left: "Odd" } });
  });

  it("reads another producer's different first page with no first footer as a blank first footer", () => {
    // Word's "Different first page" with a first header of its own and none
    // for the footer: page one has no footer in Word, and the model says so
    // rather than repeating the default there.
    const p = write(fpDoc({ header: { left: "Rest" }, firstHeader: { left: "First" }, footer: { center: "{page}" } }));
    p["word/document.xml"] = p["word/document.xml"]!.replace('<w:footerReference w:type="first" r:id="rIdFooter1"/>', "");

    expect((Agent.read(rezip(p)) as Any).page).toEqual({
      header: { left: "Rest" },
      footer: { center: "{page}" },
      firstHeader: { left: "First" },
      firstFooter: { blocks: [] },
    });
  });

  it("skips a variant that names the same part as the default", () => {
    const p = write(fpDoc({ header: { left: "Rest" }, firstHeader: { left: "First" }, footer: { center: "{page}" } }));

    // footer:first names footer1.xml, the default's part: no firstFooter.
    expect(Object.keys((Agent.read(rezip(p)) as Any).page as Any)).toEqual(["header", "footer", "firstHeader"]);
  });
});

describe("the model rules for first and even bands", () => {
  it("publishes the first and even bands in the schema, each the shape of page.header", () => {
    const page = (((Agent.jsonSchema().properties as Any).page as Any).properties as Any) ?? {};

    // header's own band, or the blank {"blocks": []} a variant uses to empty its page.
    for (const key of ["firstHeader", "firstFooter", "evenHeader", "evenFooter"]) {
      expect(page).toHaveProperty([key]);
      expect((page[key] as Any).oneOf).toEqual([{ $ref: (page.header as Any).$ref }, { $ref: "#/definitions/blankBand" }]);
    }
    expect(Schema.PAGE_BANDS).toEqual(["header", "footer", "firstHeader", "firstFooter", "evenHeader", "evenFooter"]);
  });

  it("validates a first or even band as it validates page.header", () => {
    const paths = Agent.validate(fpDoc({ firstFooter: { left: 5 }, evenHeader: "text" })).map((e) => e.path);

    expect(paths).toContain("/page/firstFooter/left");
    expect(paths).toContain("/page/evenHeader");
  });

  it.each(["firstHeader", "firstFooter", "evenHeader", "evenFooter"])("refuses a footnote in %s, as in every band", (key) => {
    const doc = fpDoc({
      [key]: {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "Head" }, { text: "", footnote: { blocks: [{ type: "paragraph", runs: [{ text: "n" }] }] } }],
          },
        ],
      },
    });

    expect(Agent.validate(doc)).toContainEqual({ path: `/page/${key}/blocks/0/runs/1/footnote`, message: "A header or footer cannot hold a footnote." });
    const repaired = Agent.validateAndRepair(doc);
    expect((((repaired.schema.page as Any)[key] as Any).blocks as Any[])[0]).toMatchObject({ runs: [{ text: "Head" }] });
  });
});

describe("templateFirstPageOnly (not applicable in this engine)", () => {
  it("changes nothing: this engine carries no template header or footer for it to move", () => {
    // PHP's option moves the header/footer `templateParts` carries to the first
    // page. This engine takes a template's styles and theme only, so there is
    // nothing to move; true, false and absent all write the same bytes.
    const template = Agent.toBytes({ page: { header: { left: "Letterhead" } }, blocks: [] } as never);
    const plain = Agent.toBytes(fpDoc() as never, { template });

    expect(Agent.toBytes(fpDoc() as never, { template, templateFirstPageOnly: true } as never)).toEqual(plain);
    expect(Agent.toBytes(fpDoc() as never, { template, templateFirstPageOnly: false } as never)).toEqual(plain);
  });
});
