import { describe, expect, it } from "vitest";
import { Agent } from "../src";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/**
 * The RTF reader's rules one at a time, on RTF small enough to read. Mirrors PHP
 * `RtfReaderTest` case for case.
 *
 * RTF is written with String.raw so a backslash in the source is a backslash in
 * the document. `u(8212)` spells the control word for U+2014.
 */
const BS = String.fromCharCode(92);
const u = (n: number): string => BS + "u" + n;
const bytes = (s: string): Uint8Array => Uint8Array.from(s, (c) => c.charCodeAt(0));

function rtfBlocks(body: string, header = ""): Any[] {
  return (Agent.read(bytes(String.raw`{\rtf1\ansi` + header + " " + body + "}")) as Any).blocks;
}

describe("text", () => {
  it("decodes \\'hh in the Windows-1252 code page by default", () => {
    expect(rtfBlocks(String.raw`\'93Hi\'94 \'80\par`)[0].runs[0].text).toBe("“Hi” €");
  });

  it("decodes \\'hh in the code page \\ansicpg declares", () => {
    // "Привет" in Windows-1251.
    expect(rtfBlocks(String.raw`\'cf\'f0\'e8\'e2\'e5\'f2\par`, String.raw`\ansicpg1251`)[0].runs[0].text).toBe("Привет");
  });

  it("replaces a double-byte character with ONE U+FFFD, not one per byte", () => {
    expect(rtfBlocks(String.raw`a\'82\'a0b\par`, String.raw`\ansicpg932`)[0].runs[0].text).toBe(
      "a" + String.fromCharCode(0xfffd) + "b",
    );
  });

  it("reads \\uN and skips its fallback by \\ucN", () => {
    expect(rtfBlocks(String.raw`\uc1 caf${u(233)}\'e9\par`)[0].runs[0].text).toBe("café");
    expect(rtfBlocks(String.raw`\uc2 x${u(8212)}--y\par`)[0].runs[0].text).toBe("x—y");
    expect(rtfBlocks(String.raw`\uc0 x${u(8212)} y\par`)[0].runs[0].text).toBe("x—y");
  });

  it("scopes \\uc to its group", () => {
    expect(rtfBlocks(String.raw`{\uc2 a${u(8212)}??}b${u(8212)}?c\par`)[0].runs[0].text).toBe("a—b—c");
  });

  it("joins a surrogate pair written as two negative \\u values", () => {
    expect(rtfBlocks(String.raw`\uc1${u(-10180)}?${u(-8311)}?\par`)[0].runs[0].text).toBe("🎉");
  });

  it("replaces half a surrogate pair", () => {
    expect(rtfBlocks(String.raw`\uc0${u(-10180)} x\par`)[0].runs[0].text).toBe(String.fromCharCode(0xfffd) + "x");
  });

  it("writes line breaks, tabs and escaped braces as text", () => {
    expect(rtfBlocks(String.raw`a\line b\tab c \{d\}\\\par`)[0].runs[0].text).toBe("a\nb\tc {d}" + BS);
  });

  it("skips destinations that are not body text", () => {
    const blocks = rtfBlocks(
      String.raw`{\fonttbl{\f0 Arial;}}{\colortbl;\red0\green0\blue0;}{\*\generator Hand;}{\*\unknowndest secret}{\header head}{\footnote note}body\par`,
    );
    expect(blocks).toEqual([{ type: "paragraph", runs: [{ text: "body" }] }]);
  });

  it("reads the title from the info group", () => {
    expect((Agent.read(bytes(String.raw`{\rtf1{\info{\title Annual Plan}{\author Someone}}Body\par}`)) as Any).title).toBe(
      "Annual Plan",
    );
  });
});

describe("formatting", () => {
  it("scopes bold, italic, underline and strike to their group", () => {
    expect(rtfBlocks(String.raw`a{\b b{\i c}}{\ul d}{\strike e}\b0 f\par`)[0].runs).toEqual([
      { text: "a" },
      { text: "b", bold: true },
      { text: "c", bold: true, italic: true },
      { text: "d", underline: true },
      { text: "e", strike: true },
      { text: "f" },
    ]);
  });

  it("does not read \\ulc (an underline COLOUR) as underlining", () => {
    expect(rtfBlocks(String.raw`{\ulc2 plain}{\uldb under}{\ul\ulnone also}\par`)[0].runs).toEqual([
      { text: "plain" },
      { text: "under", underline: true },
      { text: "also" },
    ]);
  });

  it("subtracts what a character style sets, as Word repeats it inline", () => {
    const sheet = String.raw`{\stylesheet{\s0 Normal;}{\*\cs15\b Strong;}}`;
    expect(rtfBlocks(sheet + String.raw`x{\cs15\b y}{\b z}\par`)[0].runs).toEqual([{ text: "xy" }, { text: "z", bold: true }]);
  });

  it("turns a HYPERLINK field into a link on its result", () => {
    const field = String.raw`{\field{\*\fldinst HYPERLINK "https://example.com/a" }{\fldrslt {\ul\cf2 site}}}`;
    expect(rtfBlocks("see " + field + String.raw` now\par`)[0].runs).toEqual([
      { text: "see " },
      { text: "site", link: "https://example.com/a" },
      { text: " now" },
    ]);
  });

  it("keeps a bookmark hyperlink as an anchor", () => {
    // RTF escapes the field switch's backslash: \\l, not \l (a control word).
    const field = String.raw`{\field{\*\fldinst HYPERLINK \\l "intro"}{\fldrslt top}}`;
    expect(rtfBlocks(field + String.raw`\par`)[0].runs).toEqual([{ text: "top", link: "#intro" }]);
  });

  it("keeps another field's result and drops its instruction", () => {
    expect(rtfBlocks(String.raw`page {\field{\*\fldinst PAGE}{\fldrslt 3}}\par`)[0].runs).toEqual([{ text: "page 3" }]);
  });
});

describe("structure", () => {
  it("makes a paragraph a heading by its style name, without the style's bold", () => {
    const sheet = String.raw`{\stylesheet{\s0 Normal;}{\s2\b\fs28 heading 2;}}`;
    // \pard resets paragraph properties only; \plain resets the character ones.
    expect(rtfBlocks(sheet + String.raw`\pard\plain\s2\b\fs28 Title\par\pard\plain\s0 Body\par`)).toEqual([
      { type: "heading", level: 2, runs: [{ text: "Title" }] },
      { type: "paragraph", runs: [{ text: "Body" }] },
    ]);
  });

  it("makes a paragraph a heading by its outline level", () => {
    expect(rtfBlocks(String.raw`\pard\outlinelevel0 Top\par`)[0]).toEqual({ type: "heading", level: 1, runs: [{ text: "Top" }] });
  });

  it("does not guess that a bold paragraph is a heading", () => {
    expect(rtfBlocks(String.raw`\b Just bold\b0\par`)[0]).toEqual({
      type: "paragraph",
      runs: [{ text: "Just bold", bold: true }],
    });
  });

  it("reads lists, nesting by \\ilvl and numbered by the list table", () => {
    const tables =
      String.raw`{\*\listtable{\list{\listlevel\levelnfc23{\leveltext \'01${u(8226)} ?;}}\listid10}` +
      String.raw`{\list{\listlevel\levelnfc0{\leveltext \'02\'00.;}}\listid20}}` +
      String.raw`{\*\listoverridetable{\listoverride\listid10\ls1}{\listoverride\listid20\ls2}}`;
    const body =
      String.raw`\pard\ls1\ilvl0{\listtext ${u(8226)}?\tab}One\par` +
      String.raw`\pard\ls1\ilvl1{\listtext o\tab}Inner\par` +
      String.raw`\pard\ls2\ilvl0{\listtext 1.\tab}First\par` +
      String.raw`\pard Done\par`;

    expect(rtfBlocks(tables + body)).toEqual([
      { type: "list", items: [{ runs: [{ text: "One" }], children: [{ runs: [{ text: "Inner" }] }] }] },
      { type: "list", ordered: true, items: [{ runs: [{ text: "First" }] }] },
      { type: "paragraph", runs: [{ text: "Done" }] },
    ]);
  });

  it("reads tables, with a header row where \\trhdr marks one", () => {
    const body =
      String.raw`\trowd\trhdr\cellx1000\cellx2000\pard\intbl A\cell B\cell\row` +
      String.raw`\trowd\cellx1000\cellx2000\pard\intbl 1\cell 2\cell\row` +
      String.raw`\pard After\par`;

    expect(rtfBlocks(body)).toEqual([
      {
        type: "table",
        rows: [
          { header: true, cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "A" }] }] }, { blocks: [{ type: "paragraph", runs: [{ text: "B" }] }] }] },
          { cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "1" }] }] }, { blocks: [{ type: "paragraph", runs: [{ text: "2" }] }] }] },
        ],
      },
      { type: "paragraph", runs: [{ text: "After" }] },
    ]);
  });

  it("writes a page break as its own block", () => {
    expect(rtfBlocks(String.raw`a\par\page b\par`).map((b) => b.type)).toEqual(["paragraph", "pageBreak", "paragraph"]);
  });
});
