import { describe, expect, it } from "vitest";
import cases from "../test/fixtures/charts.json";
import { Agent, DEFAULT_PALETTE, chartLabel, chartNum, chartParseNum, unzipSync, zipSync } from "../src";

/*
 * The `chart` block (last-word#4): a native Word chart written from the same
 * ECharts `option` dark-slide's chart element takes. The PHP package is the
 * reference; parity.test.ts diffs every chart part against it. These cases are
 * the shared fixture (test/fixtures/charts.json, copied verbatim from PHP).
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const CASES: Record<string, Any> = Object.fromEntries(Object.entries(cases).filter(([k]) => k !== "_comment"));
const doc = (name: string): Any => structuredClone(CASES[name]);
const dec = new TextDecoder();

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = dec.decode(data);
  return out;
}

function binaryParts(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

function patch(bytes: Uint8Array, part: string, edit: ((xml: string) => string) | null): Uint8Array {
  const files = Object.entries(unzipSync(bytes))
    .filter(([name]) => edit !== null || name !== part)
    .map(([name, data]) => ({ name, data: name === part && edit ? new TextEncoder().encode(edit(dec.decode(data))) : data }));
  return zipSync(files);
}

describe("chart block: the package", () => {
  it("writes a chart part, its relationship, its content type and an embedded workbook", () => {
    const p = parts(Agent.toBytes(doc("column")));
    expect(Object.keys(p)).toEqual(
      expect.arrayContaining(["word/charts/chart1.xml", "word/charts/_rels/chart1.xml.rels", "word/embeddings/Microsoft_Excel_Worksheet1.xlsx"]),
    );
    expect(p["[Content_Types].xml"]).toContain(
      '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>',
    );
    expect(p["[Content_Types].xml"]).toContain('<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>');
    expect(p["word/_rels/document.xml.rels"]).toMatch(
      /<Relationship Id="rId\d+" Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/chart" Target="charts\/chart1.xml"\/>/,
    );
    expect(p["word/charts/_rels/chart1.xml.rels"]).toContain('Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"');
    expect(p["word/charts/chart1.xml"]).toContain('<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>');
  });

  it("puts the chart inline, sized from widthPx / heightPx, the column by default, or a percentage", () => {
    const body = parts(Agent.toBytes(doc("column")))["word/document.xml"]!;
    expect(body).toContain('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">');
    expect(body).toContain('<wp:extent cx="5715000" cy="3048000"/>');
    expect(body).toContain('descr="Revenue by quarter, 2025 against 2026"');
    expect(body).toMatch(/<wp:docPr id="\d+" name="Chart 1"/);

    expect(parts(Agent.toBytes(doc("minimal")))["word/document.xml"]).toContain('<wp:extent cx="5943600" cy="3566160"/>');
    const half = doc("minimal");
    half.blocks[0].width = 50;
    expect(parts(Agent.toBytes(half))["word/document.xml"]).toContain('<wp:extent cx="2971800" cy="1783080"/>');
    const wide = doc("minimal");
    wide.blocks[0].widthPx = 2000;
    expect(parts(Agent.toBytes(wide))["word/document.xml"]).toContain('<wp:extent cx="5943600" ');
  });

  it("embeds a real workbook holding the plotted numbers, referenced by the series", () => {
    const outer = binaryParts(Agent.toBytes(doc("column")));
    const xlsx = parts(outer["word/embeddings/Microsoft_Excel_Worksheet1.xlsx"]!);
    expect(Object.keys(xlsx)).toEqual(["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml", "xl/styles.xml"]);
    expect(xlsx["xl/worksheets/sheet1.xml"]).toContain('<c r="B1" t="inlineStr"><is><t xml:space="preserve">2025</t></is></c>');
    expect(xlsx["xl/worksheets/sheet1.xml"]).toContain('<c r="A2" t="inlineStr"><is><t xml:space="preserve">Q1</t></is></c>');
    expect(xlsx["xl/worksheets/sheet1.xml"]).toContain('<c r="C4" s="1"><v>171.5</v></c>');
    expect(xlsx["xl/styles.xml"]).toContain('<numFmt numFmtId="164" formatCode="#,##0"/>');

    const chart = dec.decode(outer["word/charts/chart1.xml"]!);
    expect(chart).toContain("<c:f>Sheet1!$B$1</c:f>");
    expect(chart).toContain("<c:f>Sheet1!$A$2:$A$5</c:f>");
    expect(chart).toContain("<c:f>Sheet1!$C$2:$C$5</c:f>");
    expect(chart).toContain("<c:formatCode>#,##0</c:formatCode>");
  });

  it("writes identical bytes twice", () => {
    expect(Agent.toBytes(doc("column"))).toEqual(Agent.toBytes(doc("column")));
  });
});

describe("chart block: kinds", () => {
  const chart = (name: string, mutate?: (d: Any) => void): string => {
    const d = doc(name);
    mutate?.(d);
    return parts(Agent.toBytes(d))["word/charts/chart1.xml"]!;
  };

  it("column, horizontal bar and stacked bar", () => {
    expect(chart("column")).toContain('<c:barDir val="col"/><c:grouping val="clustered"/>');
    const h = chart("horizontalBar");
    expect(h).toContain('<c:barDir val="bar"/>');
    expect(h).toMatch(/<c:catAx>.*<c:axPos val="l"\/>.*<a:t>Region<\/a:t>/);
    expect(h).toMatch(/<c:valAx>.*<c:axPos val="b"\/>.*<a:t>People<\/a:t>/);
    expect(h).toContain('<c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="DC2626"/>');
    const s = chart("stackedColumn");
    expect(s).toContain('<c:grouping val="stacked"/>');
    expect(s).toContain('<c:overlap val="100"/>');
    expect(s).toContain('<c:dLblPos val="ctr"/>');
  });

  it("line and pie", () => {
    const l = chart("line");
    expect(l).toContain('<c:smooth val="1"/>');
    expect(l).toContain('<c:numFmt formatCode="0.0%" sourceLinked="0"/>');
    const p = chart("pie");
    expect(p.match(/<c:dPt>/g)?.length).toBe(4);
    expect(p).toContain('<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="0F766E"/>');
    expect(p).toContain('<c:dLblPos val="bestFit"/>');
    expect(p).not.toContain("<c:catAx>");
  });

  it("title, legend and labels only when asked; palette precedence", () => {
    const bare = chart("minimal");
    expect(bare).toContain('<c:autoTitleDeleted val="1"/>');
    expect(bare).not.toContain("<c:legend>");
    expect(bare).toContain("<c:v>Series 1</c:v>");
    expect(chart("column")).toContain('<c:roundedCorners val="0"/>');
    expect(chart("column")).toContain('<c:min val="0"/>');

    const pal = chart("stackedColumn", (d) => (d.blocks[0].option.color = ["#112233", "#445566"]));
    expect(pal).toContain('<a:srgbClr val="112233"/>');
    expect(pal).toContain('<a:srgbClr val="445566"/>');
    expect(chart("column")).toContain(`<a:srgbClr val="${DEFAULT_PALETTE[0]}"/>`);
  });
});

describe("chart block: numbers", () => {
  it("spells numbers like PHP: shortest round-trip, never an exponent", () => {
    expect(chartNum(120)).toBe("120");
    expect(chartNum(-0)).toBe("0");
    expect(chartNum(171.5)).toBe("171.5");
    expect(chartNum(0.1)).toBe("0.1");
    expect(chartNum(-1.25)).toBe("-1.25");
    expect(chartNum(1234567.891)).toBe("1234567.891");
    expect(chartNum(3e-7)).toBe("0.0000003");
    expect(chartNum(1e16)).toBe("10000000000000000");
    expect(chartNum(0.1 + 0.2)).toBe("0.30000000000000004");
    expect(chartNum(1e21)).toBe("1000000000000000000000");
    expect(chartNum(Infinity)).toBe("0");
  });

  it("labels and parsing follow the same rule", () => {
    expect(chartLabel(2025)).toBe("2025");
    expect(chartLabel(1e16)).toBe("10000000000000000");
    expect(chartLabel("Q1")).toBe("Q1");
    expect(chartLabel(true)).toBeNull();
    expect(chartParseNum(" 12.5 ")).toBe(12.5);
    expect(chartParseNum("3E-7")).toBe(3e-7);
    expect(chartParseNum("0x10")).toBe(0);
    expect(chartParseNum("")).toBe(0);
  });
});

describe("chart block: reading back", () => {
  for (const name of Object.keys(CASES)) {
    it(`reads ${name} back to the document that wrote it`, () => {
      expect(Agent.read(Agent.toBytes(doc(name)))).toEqual(doc(name));
    });
  }

  it("reads a percentage width back as the widthPx it resolved to", () => {
    const d = doc("minimal");
    d.blocks[0].width = 50;
    const chart: Any = Agent.read(Agent.toBytes(d)).blocks[0];
    expect(chart.widthPx).toBe(312);
    expect(chart.heightPx).toBe(187);
    expect(chart.width).toBeUndefined();
  });

  it("reads option.color back as each series' own colour", () => {
    const d = doc("stackedColumn");
    d.blocks[0].option.color = ["#112233", "#445566"];
    const series: Any = (Agent.read(Agent.toBytes(d)).blocks[0] as Any).option.series;
    expect(series[0].itemStyle).toEqual({ color: "#112233" });
    expect(series[1].itemStyle).toEqual({ color: "#445566" });
  });

  it("ignores a chart it cannot map instead of failing the read", () => {
    const d = { blocks: [{ type: "paragraph", runs: [{ text: "Before" }] }, ...doc("minimal").blocks, { type: "paragraph", runs: [{ text: "After" }] }] };
    const bytes = Agent.toBytes(d);
    const expected = [
      { type: "paragraph", runs: [{ text: "Before" }] },
      { type: "paragraph", runs: [{ text: "After" }] },
    ];
    const radar = patch(bytes, "word/charts/chart1.xml", (x) => x.replace("<c:barChart>", "<c:radarChart>").replace("</c:barChart>", "</c:radarChart>"));
    expect(Agent.read(radar).blocks).toEqual(expected);
    expect(Agent.read(patch(bytes, "word/charts/chart1.xml", () => "not xml <<<")).blocks).toEqual(expected);
    expect(Agent.read(patch(bytes, "word/charts/chart1.xml", null)).blocks).toEqual(expected);
  });
});

describe("chart block: teaching it", () => {
  for (const name of Object.keys(CASES)) {
    it(`accepts ${name}`, () => {
      expect(Agent.validate(doc(name))).toEqual([]);
    });
  }

  const bad: [string, Any, string, string][] = [
    ["no option", { type: "chart" }, "/blocks/0/option", "option"],
    ["no series", { type: "chart", option: { xAxis: { type: "category", data: ["a"] } } }, "/blocks/0/option/series", "series"],
    ["unsupported type", { type: "chart", option: { series: [{ type: "radar", data: [1] }] } }, "/blocks/0/option/series/0/type", "bar, line, pie"],
    [
      "mixed types",
      { type: "chart", option: { series: [{ type: "bar", data: [1] }, { type: "line", data: [2] }] } },
      "/blocks/0/option/series/1/type",
      "same type",
    ],
    [
      "two pies",
      { type: "chart", option: { series: [{ type: "pie", data: [{ name: "a", value: 1 }] }, { type: "pie", data: [{ name: "b", value: 1 }] }] } },
      "/blocks/0/option/series",
      "one series",
    ],
    ["data not a list", { type: "chart", option: { series: [{ type: "bar", data: "oops" }] } }, "/blocks/0/option/series/0/data", "data"],
    ["bad point", { type: "chart", option: { series: [{ type: "bar", data: [1, "x"] }] } }, "/blocks/0/option/series/0/data/1", "number"],
    ["values instead of data", { type: "chart", option: { series: [{ type: "bar", values: [1] }] } }, "/blocks/0/option/series/0/data", "data"],
    ["bad colour", { type: "chart", option: { series: [{ type: "bar", data: [1], itemStyle: { color: "red" } }] } }, "/blocks/0/option/series/0/itemStyle/color", "#RRGGBB"],
    ["bad width", { type: "chart", width: 150, option: { series: [{ type: "bar", data: [1] }] } }, "/blocks/0/width", "percentage"],
    ["bad numberFormat", { type: "chart", numberFormat: 7, option: { series: [{ type: "bar", data: [1] }] } }, "/blocks/0/numberFormat", "number format"],
  ];
  for (const [label, block, path, needle] of bad) {
    it(`rejects ${label}, saying how to fix it`, () => {
      const errors = Agent.validate({ blocks: [block] });
      const hit = errors.find((e) => e.path === path);
      expect(hit, JSON.stringify(errors)).toBeDefined();
      expect(hit!.message).toContain(needle);
    });
  }

  it("repairs the shorthand an agent reaches for into an ECharts option", () => {
    const result = Agent.validateAndRepair({
      blocks: [{ type: "chart", chartType: "column", title: "Revenue", categories: ["Q1", "Q2"], series: [{ name: "2026", values: [10, 12], color: "#1F3A5F" }] }],
    });
    expect(result.ok).toBe(true);
    expect(result.schema.blocks[0]).toEqual({
      type: "chart",
      option: {
        title: { text: "Revenue" },
        xAxis: { type: "category", data: ["Q1", "Q2"] },
        yAxis: { type: "value" },
        series: [{ type: "bar", name: "2026", data: [10, 12], itemStyle: { color: "#1F3A5F" } }],
      },
    });

    const bar: Any = (Agent.validateAndRepair({ blocks: [{ type: "chart", chartType: "bar", categories: ["a"], series: [{ values: [1] }] }] }).schema.blocks[0] as Any).option;
    expect(bar.yAxis).toEqual({ type: "category", data: ["a"] });
    expect(bar.xAxis).toEqual({ type: "value" });

    const pie: Any = (Agent.validateAndRepair({ blocks: [{ type: "chart", chartType: "pie", categories: ["a", "b"], series: [{ name: "Mix", values: [1, 2] }] }] }).schema.blocks[0] as Any).option;
    expect(pie).toEqual({ series: [{ type: "pie", name: "Mix", data: [{ name: "a", value: 1 }, { name: "b", value: 2 }] }] });
  });

  it("repairs series.values to data, and drops a chart it cannot save", () => {
    const fixed = Agent.validateAndRepair({ blocks: [{ type: "chart", option: { xAxis: { type: "category", data: ["a"] }, series: [{ type: "bar", values: [3] }] } }] });
    expect(fixed.ok).toBe(true);
    expect((fixed.schema.blocks[0] as Any).option.series[0]).toEqual({ type: "bar", data: [3] });

    const dropped = Agent.validateAndRepair({ blocks: [{ type: "chart", option: { series: [{ type: "radar", data: [1] }] } }] });
    expect(dropped.ok).toBe(true);
    expect(dropped.schema.blocks).toEqual([]);
  });

  it("publishes the chart block in the JSON schema, with the supported types", () => {
    const schema: Any = Agent.jsonSchema();
    const block = schema.definitions.block;
    expect(block.properties.type.enum).toContain("chart");
    expect(block.properties.option.properties.series.items.properties.type.enum).toEqual(["bar", "line", "pie"]);
    expect(block.properties).toHaveProperty("numberFormat");
  });

  it("teaches the chart block in the agent guide", () => {
    expect(Agent.guide()).toContain('{type:"chart"');
  });

  it("renders a chart into markdown as its data table", () => {
    const md = Agent.toMarkdown(doc("column"));
    expect(md).toContain("**Revenue by quarter**\n\n| Quarter | 2025 | 2026 |\n| --- | --- | --- |\n| Q1 | 120 | 130 |");
    expect(md).toContain("| Q3 | 150 | 171.5 |");
  });
});
