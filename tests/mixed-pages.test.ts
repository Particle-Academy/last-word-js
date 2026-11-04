import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { Agent, Schema, unzipSync, zipSync } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/*
 * Mixed page orientation, stage 2 (.ai/plans/mixed-page-orientation.md): a pageBreak
 * with `page` starts a new Word section. Case for case with PHP's MixedPagesTest, less
 * its template, `{blocks}` band and border cases, which this engine does not write;
 * every part held byte-identical to the PHP writer.
 */

const para = (text = "Body"): Any => ({ type: "paragraph", runs: [{ text }] });
const dec = new TextDecoder();
const enc = new TextEncoder();
const parts = (bytes: Uint8Array): Record<string, Uint8Array> => unzipSync(bytes) as Record<string, Uint8Array>;
const part = (bytes: Uint8Array, name: string): string => dec.decode(parts(bytes)[name]!);
const sections = (bytes: Uint8Array): string[] => part(bytes, "word/document.xml").match(/<w:sectPr>.*?<\/w:sectPr>|<w:sectPr\/>/gs) ?? [];

function refs(sectPr: string, bytes: Uint8Array, kind: string): Record<string, string> {
  const rels = part(bytes, "word/_rels/document.xml.rels");
  const targets = Object.fromEntries([...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
  return Object.fromEntries([...sectPr.matchAll(new RegExp(`<w:${kind}Reference w:type="(\\w+)" r:id="([^"]+)"/>`, "g"))].map((m) => [m[1], targets[m[2]!]!]));
}

function docx(body: string, extra: Record<string, string> = {}): Uint8Array {
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`;
  return zipSync([{ name: "word/document.xml", data: enc.encode(xml) }, ...Object.entries(extra).map(([name, v]) => ({ name, data: enc.encode(v) }))]);
}

const P = "<w:p><w:r><w:t>Text</w:t></w:r></w:p>";
const PORTRAIT = '<w:pgSz w:w="12240" w:h="15840"/>';
const LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>';
const ends = (inner: string): string => `<w:p><w:pPr><w:sectPr>${inner}</w:sectPr></w:pPr></w:p>`;

const mixed = (): Any => ({
  title: "Mixed",
  page: {
    margins: { top: 54 },
    header: { left: "Report", right: "Page {page} of {pages}" },
    footer: { center: "Section {section}, page of {sectionpages}" },
  },
  blocks: [
    para("Portrait"),
    { type: "pageBreak", page: { orientation: "landscape", margins: { left: 28.8, right: 28.8 } } },
    { type: "table", rows: [{ cells: [{ blocks: [para("A")] }, { blocks: [para("B")] }] }] },
    { type: "pageBreak", page: { orientation: "portrait" } },
    para("Portrait again"),
  ],
});

describe("the vocabulary", () => {
  it.each([[{ orientation: "landscape", margins: { left: 28.8 } }], [{}], [{ header: { left: "x" } }]])("accepts a break with page %#", (page) => {
    expect(Agent.validate({ blocks: [para(), { type: "pageBreak", page }, para()] } as Any)).toEqual([]);
  });

  it("declares page on the schema, without background", () => {
    const page = (Schema.jsonSchema() as Any).definitions.block.properties.page;
    expect(page.additionalProperties).toBe(false);
    expect(Object.keys(page.properties)).toEqual(expect.arrayContaining(["size", "orientation", "margins", "header", "footer", "firstHeader", "firstFooter", "evenHeader", "evenFooter"]));
    expect(page.properties.background).toBeUndefined();
  });

  it("refuses a background on a break, in PHP's words", () => {
    expect(Agent.validate({ blocks: [{ type: "pageBreak", page: { background: "#FFFFFF" } }] } as Any).map((e: Any) => e.message)).toEqual([
      "a document has one page colour: set background on the top-level page, not on a break",
    ]);
  });

  it("refuses page on a nested break, in PHP's words", () => {
    const nested = { type: "pageBreak", page: { orientation: "landscape" } };
    const errors = Agent.validate({ blocks: [{ type: "quote", blocks: [nested] }, { type: "table", rows: [{ cells: [{ blocks: [nested] }] }] }] } as Any);
    expect(errors.map((e: Any) => e.message)).toEqual([
      "page settings go on a top-level page break; a break nested in a quote or a table cell cannot change the page",
      "page settings go on a top-level page break; a break nested in a quote or a table cell cannot change the page",
    ]);
  });

  it.each([["landscape"], [[1, 2]], [3]])("refuses a page that is not an object %#", (page) => {
    expect(Agent.validate({ blocks: [{ type: "pageBreak", page }] } as Any).map((e: Any) => e.message)).toEqual([
      "page on a break must be an object of page settings, {} for a new stretch with nothing changed",
    ]);
  });

  it("judges a break's settings with the page rules", () => {
    expect(Agent.validate({ blocks: [{ type: "pageBreak", page: { size: "a3" } }] } as Any).map((e: Any) => e.path)).toEqual(["/blocks/0/page/size"]);
  });

  it("repair moves a near-miss key into page, with a note", () => {
    const r: Any = Agent.validateAndRepair({ blocks: [{ type: "pageBreak", orientation: "landscape", page: { size: "a4" } }] } as Any);
    expect(r.ok).toBe(true);
    expect(r.schema.blocks[0]).toEqual({ type: "pageBreak", page: { size: "a4", orientation: "landscape" } });
    expect(r.errors.map((e: Any) => e.message)).toContain(`moved orientation into the break's "page": a break's page settings go there`);
  });

  it("repair says so when page already holds the key", () => {
    const r: Any = Agent.validateAndRepair({ blocks: [{ type: "pageBreak", orientation: "landscape", page: { orientation: "portrait" } }] } as Any);
    expect(r.schema.blocks[0]).toEqual({ type: "pageBreak", page: { orientation: "portrait" } });
    expect(r.errors.map((e: Any) => e.message)).toContain("dropped orientation from the break: its page already says what it is");
  });

  it("an unrelated repair cannot lose page", () => {
    const page = { orientation: "landscape", margins: { left: 28.8 } };
    const r: Any = Agent.validateAndRepair({ blocks: ["needs repair", { type: "pageBreak", page }, para()] } as Any);
    expect(r.ok).toBe(true);
    expect(r.schema.blocks[1]).toEqual({ type: "pageBreak", page });
  });

  it("repair removes a background and a nested page, with a note", () => {
    let r: Any = Agent.validateAndRepair({ blocks: [{ type: "pageBreak", background: "#FFFFFF" }] } as Any);
    expect(r.ok).toBe(true);
    expect(r.schema.blocks[0].page?.background).toBeUndefined();
    r = Agent.validateAndRepair({ blocks: [{ type: "quote", blocks: [{ type: "pageBreak", page: {} }] }] } as Any);
    expect(r.ok).toBe(true);
    expect(r.schema.blocks[0].blocks[0]).toEqual({ type: "pageBreak" });
  });
});

describe("stretches", () => {
  it("merges margins side by side and replaces every other key whole", () => {
    const s: Any = Schema.stretches({
      page: { margins: { top: 10, left: 20 }, header: { left: "a", right: "b" } },
      blocks: [{ type: "pageBreak", page: { margins: { left: 30 }, header: { right: "c" } } }],
    });
    expect(s[1].doc.page.margins).toEqual({ top: 10, left: 30 });
    expect(s[1].doc.page.header).toEqual({ right: "c" });
  });

  it("never inherits the first bands", () => {
    const s: Any = Schema.stretches({
      page: { firstHeader: { left: "cover" }, firstFooter: { left: "cf" }, header: { left: "run" }, evenHeader: { left: "even" } },
      blocks: [{ type: "pageBreak", page: {} }, { type: "pageBreak", page: { firstHeader: { left: "second" } } }, { type: "pageBreak", page: {} }],
    });
    expect(Object.keys(s[1].doc.page)).toEqual(["header", "evenHeader"]);
    expect(s[2].doc.page.firstHeader).toEqual({ left: "second" });
    expect(s[3].doc.page.firstHeader).toBeUndefined();
  });

  it("a plain or nested break starts no stretch", () => {
    const s = Schema.stretches({ blocks: [para(), { type: "pageBreak" }, { type: "quote", blocks: [{ type: "pageBreak", page: {} }] }] });
    expect(s).toHaveLength(1);
    expect(s[0]!.blocks).toHaveLength(3);
  });
});

describe("the writer", () => {
  it("puts each section's sectPr in a dedicated empty paragraph", () => {
    const bytes = Agent.toBytes(mixed());
    const xml = part(bytes, "word/document.xml");
    const s = sections(bytes);
    const spacing = '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr>';
    expect(s).toHaveLength(3);
    expect(s[0]).toContain('<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1440" w:bottom="1440" w:left="1440"');
    expect(s[1]).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1080" w:right="576" w:bottom="1440" w:left="576"');
    expect(s[2]).toContain('<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="576" w:bottom="1440" w:left="576"');
    expect(xml).toContain(`Portrait</w:t></w:r></w:p><w:p><w:pPr>${spacing}${s[0]}</w:pPr></w:p>`);
    expect(xml).toContain(`</w:tbl><w:p><w:pPr>${spacing}${s[1]}</w:pPr></w:p>`);
    expect(xml).toContain(`Portrait again</w:t></w:r></w:p>${s[2]}</w:body>`);
    expect(xml).not.toContain('w:type="page"');
    expect(s.join("")).not.toContain("<w:type");
  });

  it("gives each width its own band part", () => {
    const bytes = Agent.toBytes(mixed());
    const s = sections(bytes);
    expect(refs(s[0]!, bytes, "header")).toEqual({ default: "header1.xml" });
    expect(refs(s[1]!, bytes, "header")).toEqual({ default: "header5.xml" });
    expect(refs(s[2]!, bytes, "header")).toEqual({ default: "header6.xml" });
    expect(part(bytes, "word/header1.xml")).toContain('<w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/>');
    expect(part(bytes, "word/header5.xml")).toContain('<w:tab w:val="center" w:pos="7344"/><w:tab w:val="right" w:pos="14688"/>');
    expect(part(bytes, "word/header6.xml")).toContain('<w:tab w:val="center" w:pos="5544"/><w:tab w:val="right" w:pos="11088"/>');
  });

  it("the first stretch keeps its parts' bytes", () => {
    const page = { header: { left: "L", right: "Page {page}" }, firstFooter: { center: "F" }, evenHeader: { left: "E" } };
    const alone = Agent.toBytes({ page, blocks: [para()] } as Any);
    const both = Agent.toBytes({ page, blocks: [para(), { type: "pageBreak", page: { orientation: "landscape" } }, para()] } as Any);
    for (const name of ["word/header1.xml", "word/header3.xml", "word/footer2.xml"]) expect(part(both, name)).toBe(part(alone, name));
  });

  it("titlePg is per section", () => {
    const s = sections(
      Agent.toBytes({
        page: { header: { left: "run" }, firstHeader: { left: "cover" } },
        blocks: [para(), { type: "pageBreak", page: { orientation: "landscape" } }, para(), { type: "pageBreak", page: { firstFooter: { left: "x" } } }, para()],
      } as Any),
    );
    expect(s[0]).toContain("<w:titlePg/>");
    expect(s[1]).not.toContain("<w:titlePg/>");
    expect(s[1]).not.toContain('w:type="first"');
    expect(s[2]).toContain("<w:titlePg/>");
  });

  it("even and odd is the document's", () => {
    const bytes = Agent.toBytes({ page: { header: { left: "run" } }, blocks: [para(), { type: "pageBreak", page: { evenHeader: { left: "even" } } }, para()] } as Any);
    const s = sections(bytes);
    expect(part(bytes, "word/settings.xml")).toContain("<w:evenAndOddHeaders/>");
    expect(refs(s[0]!, bytes, "header")).toEqual({ default: "header1.xml", even: "header1.xml" });
    expect(refs(s[1]!, bytes, "header").even).toBe("header5.xml");
  });

  it.each([[{ header: { blocks: [] } }], [{ header: { left: "" } }]])("blanks a band a break empties %#", (page) => {
    const bytes = Agent.toBytes({ page: { header: { left: "run" } }, blocks: [para(), { type: "pageBreak", page }, para()] } as Any);
    const target = refs(sections(bytes)[1]!, bytes, "header").default!;
    expect(target).not.toBe("header1.xml");
    expect(part(bytes, `word/${target}`)).toContain("<w:p/>");
  });

  it("caches a section number in its own part", () => {
    const bytes = Agent.toBytes({ page: { footer: { left: "{section}" } }, blocks: [para(), { type: "pageBreak", page: {} }, para()] } as Any);
    expect(part(bytes, "word/footer1.xml")).toContain("<w:t>1</w:t>");
    expect(part(bytes, "word/footer5.xml")).toContain("<w:t>2</w:t>");
  });

  it("an empty last stretch has a paragraph", () => {
    expect(part(Agent.toBytes({ blocks: [para(), { type: "pageBreak", page: {} }] } as Any), "word/document.xml")).toMatch(/<\/w:pPr><\/w:p><w:p\/><w:sectPr>.*<\/w:sectPr><\/w:body>/s);
  });

  it("a chart spans its own column and reads back without a size", () => {
    const chart = { type: "chart", option: { xAxis: { type: "category", data: ["a"] }, yAxis: { type: "value" }, series: [{ type: "bar", data: [1] }] } };
    const read: Any = Agent.read(Agent.toBytes({ blocks: [chart, { type: "pageBreak", page: { orientation: "landscape" } }, chart] } as Any));
    expect(read.blocks.map((b: Any) => b.type)).toEqual(["chart", "pageBreak", "chart"]);
    expect(read.blocks[0].widthPx).toBeUndefined();
    expect(read.blocks[2].widthPx).toBeUndefined();
  });
});

describe("the reader", () => {
  const mar = (right: number): string => `<w:pgMar w:top="1440" w:right="${right}" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;

  it("top-level page is the first section", () => {
    const doc: Any = Agent.read(docx(P + ends(PORTRAIT + mar(720)) + P + ends(`<w:type w:val="nextPage"/>${LANDSCAPE}${mar(720)}`) + P + `<w:sectPr>${PORTRAIT}${mar(1440)}</w:sectPr>`));
    expect(doc.page).toEqual({ margins: { right: 36 } });
    expect(doc.blocks[1]).toEqual({ type: "pageBreak", page: { orientation: "landscape" } });
    expect(doc.blocks[3]).toEqual({ type: "pageBreak", page: { orientation: "portrait", margins: { right: 72 } } });
    expect(doc.unrepresented).toBeUndefined();
  });

  it("w:type is how this section starts", () => {
    const doc: Any = Agent.read(docx(P + ends(`<w:type w:val="continuous"/>${PORTRAIT}`) + P + `<w:sectPr>${LANDSCAPE}</w:sectPr>`));
    expect(doc.blocks[1]).toEqual({ type: "pageBreak", page: { orientation: "landscape" } });
    expect(doc.unrepresented).toBeUndefined();
  });

  const NOTE = { where: "/blocks/1", kind: "layout-not-preserved", count: 1 };
  const MARGINS = '<w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
  it.each([
    ["nextPage", LANDSCAPE, [{ type: "pageBreak", page: { orientation: "landscape" } }], []],
    ["nextPage", PORTRAIT, [{ type: "pageBreak", page: {} }], []],
    ["oddPage", PORTRAIT, [{ type: "pageBreak", page: {} }], [NOTE]],
    ["evenPage", LANDSCAPE, [{ type: "pageBreak", page: { orientation: "landscape" } }], [NOTE]],
    ["continuous", PORTRAIT, [], [NOTE]],
    ["continuous", LANDSCAPE, [{ type: "pageBreak", page: { orientation: "landscape" } }], []],
    ["continuous", '<w:pgSz w:w="11906" w:h="16838"/>', [{ type: "pageBreak", page: { size: "a4" } }], []],
    ["continuous", PORTRAIT + MARGINS, [], [NOTE]],
    ["continuous", `${PORTRAIT}<w:cols w:num="2"/>`, [], [NOTE]],
  ])("maps a %s start %#", (type, following, breaks, notes) => {
    const doc: Any = Agent.read(docx(P + ends(PORTRAIT) + P + `<w:sectPr><w:type w:val="${type}"/>${following}</w:sectPr>`));
    expect(doc.blocks.filter((b: Any) => b.type === "pageBreak")).toEqual(breaks);
    expect(doc.unrepresented ?? []).toEqual(notes);
  });

  it.each([`${PORTRAIT}<w:cols w:num="3"/>`, `${PORTRAIT}<w:pgNumType w:start="5"/>`, '<w:pgSz w:w="10000" w:h="14000"/>'])("reports what it drops: %s", (inner) => {
    const doc: Any = Agent.read(docx(`${P}<w:sectPr>${inner}</w:sectPr>`));
    expect(doc.page).toBeUndefined();
    expect(doc.unrepresented).toEqual([{ where: "/page", kind: "layout-not-preserved", count: 1 }]);
  });

  it.each(['<w:cols w:num="1"/>', '<w:pgNumType w:fmt="decimal"/>', '<w:pgNumType w:start="1"/>'])("says nothing of what changes nothing: %s", (inner) => {
    expect((Agent.read(docx(`${P}<w:sectPr>${PORTRAIT}${inner}</w:sectPr>`)) as Any).unrepresented).toBeUndefined();
  });

  it("a section break in a quote is reported, not read", () => {
    const sdt = `<w:sdt><w:sdtPr><w:tag w:val="lastword:quote"/></w:sdtPr><w:sdtContent><w:p><w:pPr><w:sectPr>${LANDSCAPE}</w:sectPr></w:pPr><w:r><w:t>Q</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
    const doc: Any = Agent.read(docx(`${sdt}${P}<w:sectPr>${PORTRAIT}</w:sectPr>`));
    expect(doc.blocks.map((b: Any) => b.type)).toEqual(["quote", "paragraph"]);
    expect(doc.unrepresented).toEqual([{ where: "/page", kind: "layout-not-preserved", count: 1 }]);
  });

  it("a section break falls after its paragraph, and through a wrapper", () => {
    let doc: Any = Agent.read(docx(`<w:p><w:pPr><w:sectPr>${PORTRAIT}</w:sectPr></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>${P}<w:sectPr>${LANDSCAPE}</w:sectPr>`));
    expect(doc.blocks.map((b: Any) => b.type)).toEqual(["paragraph", "pageBreak", "paragraph"]);
    doc = Agent.read(docx(`<w:sdt><w:sdtContent>${P}${ends(PORTRAIT)}</w:sdtContent></w:sdt>${P}<w:sectPr>${LANDSCAPE}</w:sectPr>`));
    expect(doc.blocks.map((b: Any) => b.type)).toEqual(["paragraph", "pageBreak", "paragraph"]);
  });

  it("Word's band inheritance reads as no change", () => {
    const hdr = (t: string): string => `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:hdr>`;
    const rels =
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
      '<Relationship Id="rH2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/></Relationships>';
    const doc: Any = Agent.read(
      docx(P + ends(`<w:headerReference w:type="default" r:id="rH1"/>${PORTRAIT}`) + P + ends(LANDSCAPE) + P + `<w:sectPr><w:headerReference w:type="default" r:id="rH2"/>${PORTRAIT}</w:sectPr>`, {
        "word/header1.xml": hdr("One"),
        "word/header2.xml": hdr("Two"),
        "word/_rels/document.xml.rels": rels,
      }),
    );
    expect(doc.page.header).toEqual({ left: "One" });
    expect(doc.blocks[1]).toEqual({ type: "pageBreak", page: { orientation: "landscape" } });
    expect(doc.blocks[3]).toEqual({ type: "pageBreak", page: { orientation: "portrait", header: { left: "Two" } } });
  });

  const FIXPOINT: Record<string, Any> = {
    mixed: mixed(),
    "empty break": { blocks: [para(), { type: "pageBreak", page: {} }, para()] },
    bands: {
      page: { header: { left: "run" }, firstHeader: { left: "cover" }, evenFooter: { right: "{page}" }, footer: { left: "f" } },
      blocks: [para(), { type: "pageBreak", page: { orientation: "landscape", firstFooter: { left: "second" } } }, para(), { type: "pageBreak", page: { header: { blocks: [] }, evenFooter: { blocks: [] } } }, para()],
    },
    sizes: { page: { size: "a4" }, blocks: [para(), { type: "pageBreak", page: { size: "letter", margins: { left: 36, header: 18 } } }, para(), { type: "pageBreak", page: { margins: { left: 72 } } }, para()] },
    "empty first and last stretches": { blocks: [{ type: "pageBreak", page: {} }, para(), { type: "pageBreak", page: { orientation: "landscape" } }] },
  };

  it.each(Object.keys(FIXPOINT))("%s: reads what the writer wrote, to a fixpoint", (name) => {
    const doc = FIXPOINT[name];
    const once: Any = Agent.read(Agent.toBytes(doc));
    expect(Agent.read(Agent.toBytes(once))).toEqual(once);
    expect(Agent.validate(once)).toEqual([]);
    expect(once.unrepresented).toBeUndefined();
    expect(once.blocks.filter((b: Any) => b.type === "pageBreak")).toHaveLength(doc.blocks.filter((b: Any) => b.type === "pageBreak").length);
  });

  it("reads back the model it was given, key for key", () => {
    const doc = mixed();
    const read: Any = Agent.read(Agent.toBytes(doc));
    expect(read.page).toEqual(doc.page);
    expect(read.blocks[1]).toEqual(doc.blocks[1]);
    expect(read.blocks[3]).toEqual(doc.blocks[3]);
  });

  it("reads the real two-section file as PHP does", () => {
    const doc = Agent.read(new Uint8Array(readFileSync(new URL("../test/fixtures/formats/two-sections.docx", import.meta.url)))) as Any;
    expect(doc).toEqual(JSON.parse(readFileSync(new URL("../test/fixtures/formats/two-sections.read.json", import.meta.url), "utf8")));
    expect(doc.blocks[2]).toEqual({ type: "pageBreak", page: { orientation: "landscape" } });
  });

  it("describe counts the stretches", () => {
    expect(Agent.describe(mixed())).toContain("Stretches: 3 (portrait, landscape, portrait)");
    expect(Agent.describe({ blocks: [para()] } as Any)).not.toContain("Stretches");
  });

  /* ---------------- held to PHP, part for part ---------------- */
  const dir = mkdtempSync(join(tmpdir(), "lw-mp-"));
  afterAll(() => rmSync(dir, { recursive: true, force: true }));
  const SCRIPT = join(__dirname, "..", "scripts", "php-tobytes.php");

  function phpBytes(doc: Any, name: string): Uint8Array {
    const safe = name.replace(/\W+/g, "-");
    const input = join(dir, `${safe}.json`);
    const out = join(dir, `${safe}.docx`);
    writeFileSync(input, JSON.stringify(doc));
    execFileSync("php", [SCRIPT, input, out], { shell: true, windowsHide: true });
    return new Uint8Array(readFileSync(out));
  }

  it.each(Object.keys(FIXPOINT))("%s: every part is the PHP writer's, and reads back as PHP's does", (name) => {
    const ours = parts(Agent.toBytes(FIXPOINT[name]));
    const theirBytes = phpBytes(FIXPOINT[name], name);
    const theirs = parts(theirBytes);
    expect(Object.keys(ours).sort()).toEqual(Object.keys(theirs).sort());
    // styles and numbering are parity.test.ts's recorded divergences, and PHP declares
    // png and jpeg unconditionally. Nothing this feature writes is among them.
    for (const key of Object.keys(ours)) {
      if (key === "word/styles.xml" || key === "word/numbering.xml") continue;
      const want = dec.decode(theirs[key]).replace(/<Default Extension="(png|jpeg)" ContentType="[^"]+"\/>/g, "");
      expect(dec.decode(ours[key]), key).toBe(want);
    }
    expect(Agent.read(theirBytes)).toEqual(Agent.read(Agent.toBytes(FIXPOINT[name])));
  });
});
