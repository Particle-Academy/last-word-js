import { describe, expect, it } from "vitest";
import { Agent, Schema, zipSync } from "../src";
import { layoutRows } from "../src/writer/docx-writer";

// colSpan has a ceiling, case for case with PHP's TableSpanLimitTest and
// Python's test_table_span_limit.py. rowSpan has none on purpose.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const spanDoc = (cell: Record<string, unknown>) => ({
  blocks: [{ type: "table", rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [{ text: "kept" }] }], ...cell }] }] }],
});

const encoder = new TextEncoder();

function docxWithGridSpan(val: string): Uint8Array {
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr>` +
    `<w:tc><w:tcPr><w:gridSpan w:val="${val}"/></w:tcPr><w:p><w:r><w:t>kept</w:t></w:r></w:p></w:tc>` +
    `</w:tr></w:tbl></w:body></w:document>`;
  return zipSync([{ name: "word/document.xml", data: encoder.encode(xml) }]);
}

describe("colSpan ceiling", () => {
  it("is 1000", () => {
    expect(Schema.MAX_COL_SPAN).toBe(1000);
  });

  it("validate accepts a colSpan at the ceiling and refuses one past it, by name", () => {
    expect(Agent.validate(spanDoc({ colSpan: 1000 }))).toEqual([]);
    expect(Agent.validate(spanDoc({ colSpan: 1001 }))).toEqual([
      { path: "/blocks/0/rows/0/cells/0/colSpan", message: "Cell `colSpan` must be at most 1000." },
    ]);
  });

  it("toBytes refuses past the ceiling before it does any work", () => {
    expect(() => Agent.toBytes(spanDoc({ colSpan: Number.MAX_SAFE_INTEGER }))).toThrow(/schema validation/);
    expect(() => Agent.toBytes(spanDoc({ colSpan: 1e308 }))).toThrow(/schema validation/);
  });

  it("rowSpan has no ceiling: a merge down a long table stays valid", () => {
    expect(Agent.validate(spanDoc({ rowSpan: 1000000 }))).toEqual([]);
    expect(Agent.toBytes(spanDoc({ rowSpan: 1000000 })).length).toBeGreaterThan(0);
  });

  it("validateAndRepair clamps, says so and keeps the text", () => {
    const out: Any = Agent.validateAndRepair(spanDoc({ colSpan: 1000000000 }));
    expect(out.ok).toBe(true);
    const cell = out.schema.blocks[0].rows[0].cells[0];
    expect(cell.colSpan).toBe(1000);
    expect(cell.blocks[0].runs[0].text).toBe("kept");
    expect(out.errors).toContainEqual({
      path: "/blocks/0/rows/0/cells/0/colSpan",
      message: "Clamped colSpan 1000000000 to 1000.",
    });
  });

  it("the JSON schema declares the ceiling and leaves rowSpan open", () => {
    const json = JSON.stringify(Agent.jsonSchema());
    expect(json).toContain('"colSpan":{"type":"integer","minimum":1,"maximum":1000}');
    expect(json).not.toContain('"rowSpan":{"type":"integer","minimum":1,"maximum"');
  });

  it("the writer clamps a document that never met the validator", () => {
    const { laid, colCount } = layoutRows(spanDoc({ colSpan: 2000 }).blocks[0]!.rows);
    expect(colCount).toBe(1000);
    expect(laid[0]![0]!.span).toBe(1000);
    expect(layoutRows(spanDoc({ colSpan: 1e308 }).blocks[0]!.rows).colCount).toBe(1000);
  });

  const reads: Array<[string, number | undefined]> = [
    ["3", 3],
    ["1000", 1000],
    ["1000000000", 1000],
    ["99999999999999999999999", 1000],
    ["1e30", 1000],
    ["1e400", undefined],
    ["-5", undefined],
    ["2.7", 2],
    ["abc", undefined],
  ];
  it.each(reads)("reading gridSpan %s is bounded and keeps the text", (val, want) => {
    const doc: Any = Agent.read(docxWithGridSpan(val));
    const cell = doc.blocks[0].rows[0].cells[0];
    expect(cell.colSpan).toBe(want);
    expect(cell.blocks[0].runs[0].text).toBe("kept");
    expect(Agent.validate(doc)).toEqual([]);
  });
});
