import { describe, it, expect } from "vitest";
import { Agent, fromMarkdown, parseInline, toMarkdown } from "../src";
import canonical from "../test/fixtures/canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

describe("markdown bridge — fixpoint (vector 2)", () => {
  it("fromMarkdown(toMarkdown(doc)) reaches a semantic fixpoint on the canonical doc", () => {
    const md1 = toMarkdown(canonical as Any);
    const doc2 = fromMarkdown(md1);
    const md2 = toMarkdown(doc2);
    expect(md2).toBe(md1);

    const doc3 = fromMarkdown(md2);
    expect(doc3).toEqual(doc2);
  });

  it("keeps GFM-expressible content across the bridge", () => {
    const md = toMarkdown(canonical as Any);
    expect(md).toContain("# LastWord Canonical");
    expect(md).toContain("**bold**");
    expect(md).toContain("*italic*");
    expect(md).toContain("~~strikethrough~~");
    expect(md).toContain("`inline code`");
    expect(md).toContain("[link](https://particle.academy)");
    expect(md).toContain("- Alpha");
    expect(md).toContain("    - *Alpha one deep*");
    expect(md).toContain("1. First");
    expect(md).toContain("| **Name** | **Format** | **Notes** |");
    expect(md).toContain("```typescript");
    expect(md).toContain("> The last word in agentic documents.");
    expect(md).toContain("![Tiny red square](data:image/png;base64,");
    expect(md).toContain("---");
  });
});

describe("markdown bridge — normalized identity (vector 3)", () => {
  it("toMarkdown(fromMarkdown(md)) === md for a canonical GFM doc", () => {
    const md =
      [
        "# LastWord Canonical",
        "",
        "Some **bold**, *italic*, ~~struck~~, `code`, and a [link](https://particle.academy).",
        "",
        "## Lists",
        "",
        "- Alpha",
        "  - Beta",
        "    - Gamma",
        "- Delta",
        "",
        "1. First",
        "   1. Nested",
        "2. Second",
        "",
        "| Name | Format |",
        "| --- | --- |",
        "| holy-sheet | xlsx |",
        "| dark-slide | pptx |",
        "",
        "```typescript",
        "const x: number = 42;",
        "```",
        "",
        "> The last word in agentic documents.",
        "",
        "![Tiny red square](data:image/png;base64,iVBORw0KGgo=)",
        "",
        "---",
        "",
        "The end.",
      ].join("\n") + "\n";

    expect(toMarkdown(fromMarkdown(md))).toBe(md);
  });

  it("parses nested mixed structures from markdown", () => {
    const doc: Any = fromMarkdown(
      ["- a", "  - b", "    - c", "", "1. one", "2. two"].join("\n"),
    );
    expect(doc.blocks[0]).toEqual({
      type: "list",
      items: [{ runs: [{ text: "a" }], children: [{ runs: [{ text: "b" }], children: [{ runs: [{ text: "c" }] }] }] }],
    });
    expect(doc.blocks[1]).toEqual({
      type: "list",
      ordered: true,
      items: [{ runs: [{ text: "one" }] }, { runs: [{ text: "two" }] }],
    });
  });

  it("parses tables with a header row", () => {
    const doc: Any = fromMarkdown(["| A | B |", "| --- | --- |", "| 1 | 2 |"].join("\n"));
    expect(doc.blocks[0].type).toBe("table");
    expect(doc.blocks[0].rows[0].header).toBe(true);
    expect(doc.blocks[0].rows[1].cells[1].blocks[0].runs).toEqual([{ text: "2" }]);
  });

  it("is exposed on the Agent façade", () => {
    const doc = Agent.fromMarkdown("# Hi\n");
    expect(Agent.toMarkdown(doc)).toBe("# Hi\n");
  });
});

describe("inline tokenizer", () => {
  it("handles bold+italic+strike nesting", () => {
    expect(parseInline("~~***x***~~")).toEqual([{ text: "x", bold: true, italic: true, strike: true }]);
  });

  it("handles code spans that swallow markers", () => {
    expect(parseInline("`**not bold**`")).toEqual([{ text: "**not bold**", code: true }]);
  });

  it("handles multi-backtick code spans containing backticks", () => {
    expect(parseInline("`` a`b ``")).toEqual([{ text: "a`b", code: true }]);
  });

  it("applies links to styled inner runs", () => {
    expect(parseInline("[**docs**](https://x.dev)")).toEqual([
      { text: "docs", bold: true, link: "https://x.dev" },
    ]);
  });

  it("honors backslash escapes", () => {
    expect(parseInline("\\*literal\\* \\[brackets\\]")).toEqual([{ text: "*literal* [brackets]" }]);
  });

  it("keeps snake_case intact", () => {
    expect(parseInline("use snake_case_names here")).toEqual([{ text: "use snake_case_names here" }]);
  });

  it("degrades inline images to their alt text", () => {
    expect(parseInline("see ![tiny](data:image/png;base64,x) here")).toEqual([{ text: "see tiny here" }]);
  });
});

describe("markdown drops what GFM cannot express (documented lossiness)", () => {
  it("drops underline/color/highlight decoration but keeps text", () => {
    const md = toMarkdown({
      blocks: [
        {
          type: "paragraph",
          runs: [
            { text: "u", underline: true },
            { text: " c", color: "#FF0000" },
            { text: " h", highlight: "#FFFF00" },
          ],
        },
      ],
    } as Any);
    expect(md).toBe("u c h\n");
  });

  it("drops pageBreak blocks", () => {
    const md = toMarkdown({ blocks: [{ type: "paragraph", runs: [{ text: "a" }] }, { type: "pageBreak" }] } as Any);
    expect(md).toBe("a\n");
  });
});
