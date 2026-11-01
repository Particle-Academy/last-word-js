import { describe, expect, it } from "vitest";
import { Agent, Schema, unzipSync, zipSync } from "../src";
import { parseXml, type XmlNode } from "../src/reader/xml";

/**
 * Word's standard fields in header and footer text, beyond `{page}` /
 * `{pages}` (PHP last-word 0.16.0, MOIC 2026-11-01), and the document metadata
 * AUTHOR and TITLE read from. Mirrors PHP's BandFieldsTest; byte parity with
 * PHP is in parity.test.ts ("fieldTokens", "meta").
 *
 * Grammar, in a band slot:
 *
 *   {page} {pages} {section} {sectionpages} {author} {title} {filename}
 *   {date} {date:FORMAT} {time} {time:FORMAT}
 *
 * FORMAT is a Word date-time picture (`yyyy-MM-dd`, `d MMMM yyyy`, `HH:mm`),
 * one or more characters, none of `{`, `}`, `"` or a backslash. Anything else in
 * braces is literal text, as it always was. In the body every token is literal.
 *
 * Each token is a COMPLEX field -- begin, ` INSTR ` instruction, separate, the
 * cached result, end -- every piece in the band's formatting, as `{page}`
 * already was. The cached result is deterministic: see FIELD_ROWS.
 *
 * (PHP also writes these in a `{blocks}` band's run text; this engine writes no
 * `{blocks}` band, so that half is not here.)
 */
type Any = Record<string, unknown>;

const BS = String.fromCharCode(92); // one backslash, spelled so no escape layer can touch it

/** token, instruction, cached result (with meta.author "Ada Lovelace" and title "Board memo"). */
const FIELD_ROWS: [string, string, string][] = [
  ["{page}", " PAGE ", "1"],
  ["{pages}", " NUMPAGES ", "1"],
  ["{section}", " SECTION ", "1"],
  ["{sectionpages}", " SECTIONPAGES ", "1"],
  ["{author}", " AUTHOR ", "Ada Lovelace"],
  ["{title}", " TITLE ", "Board memo"],
  ["{filename}", " FILENAME ", ""],
  ["{date}", " DATE ", ""],
  ["{date:yyyy-MM-dd}", ` DATE ${BS}@ "yyyy-MM-dd" `, ""],
  ["{date:d MMMM yyyy}", ` DATE ${BS}@ "d MMMM yyyy" `, ""],
  ["{time}", " TIME ", ""],
  ["{time:HH:mm}", ` TIME ${BS}@ "HH:mm" `, ""],
];

function bfDoc(page: Any, extra: Any = {}): Any {
  return {
    title: "Board memo",
    meta: { author: "Ada Lovelace" },
    page,
    blocks: [{ type: "paragraph", runs: [{ text: "Body" }] }],
    ...extra,
  };
}

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = new TextDecoder().decode(data);
  return out;
}

function rezip(p: Record<string, string>): Uint8Array {
  return zipSync(Object.entries(p).map(([name, xml]) => ({ name, data: new TextEncoder().encode(xml) })));
}

/** Every complex field in a part, in order, as [instruction, cached result]. */
function fields(xml: string): [string, string][] {
  const out: [string, string][] = [];
  let state: "instr" | "result" | null = null;
  let instr = "";
  let result = "";
  const walk = (node: XmlNode): void => {
    for (const el of node.children) {
      if (el.name === "fldChar") {
        const type = el.attrs["fldCharType"];
        if (type === "begin") [state, instr, result] = ["instr", "", ""];
        else if (type === "separate") state = "result";
        else if (type === "end") {
          out.push([instr, result]);
          state = null;
        }
      } else if (el.name === "instrText" && state === "instr") {
        instr += el.text;
      } else if (el.name === "t" && state === "result") {
        result += el.text;
      }
      walk(el);
    }
  };
  walk(parseXml(xml)!);
  return out;
}

const readPage = (bytes: Uint8Array): Any => (Agent.read(bytes) as Any).page as Any;

describe("header/footer field tokens", () => {
  it.each(FIELD_ROWS)("writes %s in a band slot as a live field with its cached result", (token, instr, result) => {
    const p = parts(Agent.toBytes(bfDoc({ footer: { left: `A ${token} B` } }) as never));

    expect(fields(p["word/footer1.xml"]!)).toEqual([[instr, result]]);
    expect(p["word/footer1.xml"]).toContain('<w:t xml:space="preserve">A </w:t>');
    expect(p["word/footer1.xml"]).toContain('<w:t xml:space="preserve"> B</w:t>');
  });

  it.each(FIELD_ROWS)("reads %s back from a slot band", (token) => {
    const page = { header: { left: `L ${token}`, right: "R" } };

    expect(readPage(Agent.toBytes(bfDoc(page) as never))).toEqual(page);
  });

  it("round-trips every token together, with meta and title", () => {
    const doc = bfDoc({ footer: { left: FIELD_ROWS.map((r) => r[0]).join(" | ") } });

    expect(Agent.read(Agent.toBytes(doc as never))).toEqual(doc);
  });

  it.each(["{Date}", "{AUTHOR}", "{foo}", "{date:}", '{date:a"b}', "{ page}", "{author:x}", "{filename:p}", `{date:a${BS}b}`])(
    "keeps %s literal",
    (text) => {
      const bytes = Agent.toBytes(bfDoc({ footer: { left: text } }) as never);

      expect(parts(bytes)["word/footer1.xml"]).not.toContain("fldChar");
      expect((readPage(bytes).footer as Any).left).toBe(text);
    },
  );

  it("keeps every token literal in the body", () => {
    const doc = { blocks: [{ type: "paragraph", runs: [{ text: "Dated {date} by {author}, {title}." }] }] };

    expect(parts(Agent.toBytes(doc as never))["word/document.xml"]).not.toContain("fldChar");
  });

  it.each([
    ["Word default date", ` DATE ${BS}@ "M/d/yyyy" ${BS}* MERGEFORMAT `, "{date:M/d/yyyy}"],
    ["unquoted picture", ` DATE ${BS}@ yyyy `, "{date:yyyy}"],
    ["lower case", " date ", "{date}"],
    ["time", ` TIME ${BS}@ "h:mm am/pm" `, "{time:h:mm am/pm}"],
    ["section pages", " SECTIONPAGES ", "{sectionpages}"],
    ["author with an argument", ' AUTHOR "Someone" ', "{author}"],
    ["filename with path", ` FILENAME ${BS}p `, "{filename}"],
    ["a picture the grammar cannot hold", ` DATE ${BS}@ "d}M" `, "{date}"],
  ])("reads another producer's field as a token: %s", (_label, instr, token) => {
    const p = parts(Agent.toBytes(bfDoc({ footer: { left: "X {date} Y" } }) as never));
    p["word/footer1.xml"] = p["word/footer1.xml"]!.replace(" DATE ", instr.replace(/&/g, "&amp;").replace(/"/g, "&quot;"));

    expect((readPage(rezip(p)).footer as Any).left).toBe(`X ${token} Y`);
  });

  it("reads a simple field (w:fldSimple) as its token too", () => {
    const p = parts(Agent.toBytes(bfDoc({ footer: { left: "X" } }) as never));
    p["word/footer1.xml"] = p["word/footer1.xml"]!.replace(
      "</w:p>",
      `<w:fldSimple w:instr=" DATE ${BS}@ &quot;yyyy&quot; "><w:r><w:t>2026</w:t></w:r></w:fldSimple></w:p>`,
    );

    // After the two tabs: the right slot.
    expect(readPage(rezip(p)).footer).toEqual({ left: "X", right: "{date:yyyy}" });
  });

  it("keeps the cached result of a field that is not a token, in a one-line band", () => {
    const p = parts(Agent.toBytes(bfDoc({ footer: { left: "X {page} Y" } }) as never));
    p["word/footer1.xml"] = p["word/footer1.xml"]!
      .replace(" PAGE ", ' STYLEREF "Heading 1" '.replace(/"/g, "&quot;"))
      .replace("<w:t>1</w:t>", "<w:t>Introduction</w:t>");

    expect((readPage(rezip(p)).footer as Any).left).toBe("X Introduction Y");
  });

  it("keeps the cached result of a simple field that is not a token", () => {
    const p = parts(Agent.toBytes(bfDoc({ footer: { left: "X" } }) as never));
    p["word/footer1.xml"] = p["word/footer1.xml"]!.replace(
      "</w:p>",
      `<w:fldSimple w:instr=" STYLEREF 1 "><w:r><w:t> Intro</w:t></w:r></w:fldSimple></w:p>`,
    );

    expect(readPage(rezip(p)).footer).toEqual({ left: "X", right: " Intro" });
  });

  it("exposes the grammar: Schema.splitFieldTokens() and Schema.fieldTokenOf()", () => {
    expect(Schema.splitFieldTokens("Page {page} of {pages}, {date:d MMMM yyyy} {Date}")).toEqual([
      "Page ",
      { token: "{page}", name: "page", instruction: "PAGE" },
      " of ",
      { token: "{pages}", name: "pages", instruction: "NUMPAGES" },
      ", ",
      { token: "{date:d MMMM yyyy}", name: "date", instruction: `DATE ${BS}@ "d MMMM yyyy"` },
      " {Date}",
    ]);
    expect(Schema.fieldTokenOf(` time ${BS}@ "HH:mm" `)).toBe("{time:HH:mm}");
    expect(Schema.fieldTokenOf(" STYLEREF 1 ")).toBeNull();
  });
});

describe("meta", () => {
  it("writes meta.author as dc:creator beside dc:title, and reads it back", () => {
    const doc = { title: "Board memo", meta: { author: "Ada & Co" }, blocks: [] };
    const p = parts(Agent.toBytes(doc as never));

    expect(p["docProps/core.xml"]).toContain(
      "<dc:title>Board memo</dc:title><dc:creator>Ada &amp; Co</dc:creator></cp:coreProperties>",
    );
    expect(Agent.read(Agent.toBytes(doc as never))).toEqual(doc);
  });

  it("writes core.xml for an author with no title", () => {
    const doc = { meta: { author: "Ada Lovelace" }, blocks: [] };
    const p = parts(Agent.toBytes(doc as never));

    expect(p["docProps/core.xml"]).toContain("<dc:creator>Ada Lovelace</dc:creator>");
    expect(p["docProps/core.xml"]).not.toContain("dc:title");
    expect(p["[Content_Types].xml"]).toContain("/docProps/core.xml");
    expect(p["_rels/.rels"]).toContain("docProps/core.xml");
    expect(Agent.read(Agent.toBytes(doc as never))).toEqual(doc);
  });

  it("writes no core part for an empty author with no title", () => {
    expect(parts(Agent.toBytes({ meta: { author: "" }, blocks: [] } as never))).not.toHaveProperty(["docProps/core.xml"]);
  });

  it("reads an empty dc:creator as no meta", () => {
    const p = parts(Agent.toBytes({ title: "T", meta: { author: "X" }, blocks: [] } as never));
    p["docProps/core.xml"] = p["docProps/core.xml"]!.replace("<dc:creator>X</dc:creator>", "<dc:creator></dc:creator>");

    expect(Agent.read(rezip(p))).toEqual({ title: "T", blocks: [] });
  });

  it("writes an empty cached result for AUTHOR and TITLE when there is no author or title", () => {
    const p = parts(Agent.toBytes({ page: { footer: { left: "{author} {title}" } }, blocks: [] } as never));

    expect(fields(p["word/footer1.xml"]!)).toEqual([
      [" AUTHOR ", ""],
      [" TITLE ", ""],
    ]);
    expect(p).not.toHaveProperty(["docProps/core.xml"]);
  });

  it("publishes meta in the schema", () => {
    const meta = (Agent.jsonSchema().properties as Any).meta as Any;

    expect(meta.type).toBe("object");
    expect(meta.additionalProperties).toBe(false);
    expect(Object.keys(meta.properties as Any)).toEqual(["author"]);
  });

  it.each([
    ["not an object", "Ada", "/meta"],
    ["author not text", { author: 7 }, "/meta/author"],
    ["title belongs at the top", { title: "Board memo" }, "/meta/title"],
  ])("validates meta: %s", (_label, meta, path) => {
    expect(Agent.validate({ meta, blocks: [] }).map((e) => e.path)).toContain(path);
  });

  it("repairs meta.title into the document title", () => {
    const out = Agent.validateAndRepair({ meta: { author: "Ada", title: "Board memo" }, blocks: [] });

    expect(out.ok).toBe(true);
    expect(out.schema.title).toBe("Board memo");
    expect(out.schema.meta).toEqual({ author: "Ada" });
  });

  it("drops meta.title when there is a title already, and drops a meta that is not an object", () => {
    const kept = Agent.validateAndRepair({ title: "Kept", meta: { title: "Lost" }, blocks: [] });
    expect(kept.ok).toBe(true);
    expect(kept.schema.title).toBe("Kept");
    expect(kept.schema.meta).toEqual({});

    const dropped = Agent.validateAndRepair({ meta: ["Ada"], blocks: [] });
    expect(dropped.ok).toBe(true);
    expect(dropped.schema).not.toHaveProperty("meta");
  });
});
