/**
 * The parts one `chart` block becomes: `word/charts/chartN.xml` (DrawingML
 * chart), its relationships, and the embedded workbook Word opens on "Edit
 * Data". Mirrors PHP `LastWord\Writer\ChartPart` byte for byte -- the
 * cross-engine parity suite diffs every one of these parts against PHP's.
 *
 * Workbook layout, as Word lays out its own: series names across row 1 from B1,
 * categories down column A from A2, one column of values per series. Every
 * series references those cells AND carries a cache of them.
 *
 * Element order follows the ECMA-376 sequences exactly: Word rejects an
 * out-of-order chart as unreadable content without naming the part.
 */
import { Xml } from "../helpers/xml";
import { chartColumn, chartNum, type ChartSeriesSpec, type ChartSpec } from "../helpers/chart";

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_SHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CAT_AX = "111111111";
const VAL_AX = "222222222";
const GRID = "D9D9D9";

export function chartXml(spec: ChartSpec, numberFormat: string | null): string {
  const plot =
    spec.kind === "line" ? lineChart(spec, numberFormat) : spec.kind === "pie" ? pieChart(spec, numberFormat) : barChart(spec, numberFormat);
  const title =
    spec.title !== ""
      ? '<c:title>' + rich(spec.title, false) + '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
      : '<c:autoTitleDeleted val="1"/>';

  return (
    Xml.declaration() +
    '<c:chartSpace xmlns:c="' + NS_C + '" xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '">' +
    '<c:date1904 val="0"/>' +
    // Absent, Word draws the chart with rounded corners.
    '<c:roundedCorners val="0"/>' +
    "<c:chart>" +
    title +
    "<c:plotArea><c:layout/>" + plot + "</c:plotArea>" +
    (spec.legend ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : "") +
    '<c:plotVisOnly val="1"/>' +
    '<c:dispBlanksAs val="gap"/>' +
    "</c:chart>" +
    // No frame: a report's chart sits in the text, not in a box.
    "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>" +
    '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' +
    "</c:chartSpace>"
  );
}

/** The chart part's relationships: its one embedded workbook. */
export function chartRelsXml(n: number): string {
  return (
    Xml.declaration() +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' +
    n +
    '.xlsx"/>' +
    "</Relationships>"
  );
}

/** The embedded workbook's parts, in archive order. */
export function workbookParts(spec: ChartSpec, numberFormat: string | null): [string, string][] {
  const rows = rowCount(spec);
  const styled = numberFormat !== null ? ' s="1"' : "";

  let sheet = '<row r="1">';
  spec.series.forEach((series, i) => {
    sheet += stringCell(chartColumn(i + 1) + "1", series.name);
  });
  sheet += "</row>";
  for (let r = 0; r < rows; r++) {
    const rowNumber = r + 2;
    sheet += '<row r="' + rowNumber + '">' + stringCell("A" + rowNumber, category(spec, r));
    spec.series.forEach((series, i) => {
      if (r < series.values.length) {
        sheet += '<c r="' + chartColumn(i + 1) + rowNumber + '"' + styled + "><v>" + chartNum(series.values[r]!) + "</v></c>";
      }
    });
    sheet += "</row>";
  }

  const numFmts =
    numberFormat !== null ? '<numFmts count="1"><numFmt numFmtId="164" formatCode="' + Xml.attr(numberFormat) + '"/></numFmts>' : "";
  const xfs =
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    (numberFormat !== null ? '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' : "");

  return [
    [
      "[Content_Types].xml",
      Xml.declaration() +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        "</Types>",
    ],
    [
      "_rels/.rels",
      Xml.declaration() +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        "</Relationships>",
    ],
    [
      "xl/workbook.xml",
      Xml.declaration() +
        '<workbook xmlns="' + NS_SHEET + '" xmlns:r="' + NS_R + '">' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
        "</workbook>",
    ],
    [
      "xl/_rels/workbook.xml.rels",
      Xml.declaration() +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        "</Relationships>",
    ],
    ["xl/worksheets/sheet1.xml", Xml.declaration() + '<worksheet xmlns="' + NS_SHEET + '"><sheetData>' + sheet + "</sheetData></worksheet>"],
    [
      "xl/styles.xml",
      Xml.declaration() +
        '<styleSheet xmlns="' + NS_SHEET + '">' +
        numFmts +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="' + (numberFormat !== null ? 2 : 1) + '">' + xfs + "</cellXfs>" +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        "</styleSheet>",
    ],
  ];
}

// ── Plots ───────────────────────────────────────────────────────────────────

function barChart(spec: ChartSpec, numberFormat: string | null): string {
  let sers = "";
  spec.series.forEach((series, idx) => {
    let dPts = "";
    for (const [i, color] of series.pointColors) {
      dPts += '<c:dPt><c:idx val="' + i + '"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>' + "<c:spPr>" + solid(color) + "</c:spPr></c:dPt>";
    }
    sers +=
      '<c:ser><c:idx val="' + idx + '"/><c:order val="' + idx + '"/>' +
      seriesName(series, idx) +
      "<c:spPr>" + solid(series.color) + "</c:spPr>" +
      '<c:invertIfNegative val="0"/>' +
      dPts +
      // outEnd is not a legal position on a stacked bar: Word refuses the file.
      dataLabels(series, spec.stacked ? "ctr" : "outEnd", numberFormat) +
      categoryRef(spec) +
      valueRef(spec, series, idx, numberFormat) +
      "</c:ser>";
  });

  return (
    "<c:barChart>" +
    '<c:barDir val="' + (spec.horizontal ? "bar" : "col") + '"/>' +
    '<c:grouping val="' + (spec.stacked ? "stacked" : "clustered") + '"/>' +
    '<c:varyColors val="0"/>' +
    sers +
    '<c:gapWidth val="150"/>' +
    (spec.stacked ? '<c:overlap val="100"/>' : "") +
    '<c:axId val="' + CAT_AX + '"/><c:axId val="' + VAL_AX + '"/>' +
    "</c:barChart>" +
    axes(spec, numberFormat, nonNegative(spec))
  );
}

function lineChart(spec: ChartSpec, numberFormat: string | null): string {
  let sers = "";
  spec.series.forEach((series, idx) => {
    let dPts = "";
    for (const [i, color] of series.pointColors) {
      dPts += '<c:dPt><c:idx val="' + i + '"/>' + marker(color) + '<c:bubble3D val="0"/></c:dPt>';
    }
    sers +=
      '<c:ser><c:idx val="' + idx + '"/><c:order val="' + idx + '"/>' +
      seriesName(series, idx) +
      '<c:spPr><a:ln w="28575" cap="rnd">' + solid(series.color) + "<a:round/></a:ln></c:spPr>" +
      marker(series.color) +
      dPts +
      dataLabels(series, "t", numberFormat) +
      categoryRef(spec) +
      valueRef(spec, series, idx, numberFormat) +
      '<c:smooth val="' + (series.smooth ? "1" : "0") + '"/>' +
      "</c:ser>";
  });

  return (
    "<c:lineChart>" +
    '<c:grouping val="standard"/>' +
    '<c:varyColors val="0"/>' +
    sers +
    '<c:marker val="1"/>' +
    '<c:axId val="' + CAT_AX + '"/><c:axId val="' + VAL_AX + '"/>' +
    "</c:lineChart>" +
    axes(spec, numberFormat, false)
  );
}

function pieChart(spec: ChartSpec, numberFormat: string | null): string {
  const series = spec.series[0]!;
  let dPts = "";
  for (const [i, color] of series.pointColors) {
    // A white hairline between slices, as Word draws them.
    dPts +=
      '<c:dPt><c:idx val="' + i + '"/><c:bubble3D val="0"/>' +
      "<c:spPr>" + solid(color) + '<a:ln w="12700">' + solid("FFFFFF") + "</a:ln></c:spPr></c:dPt>";
  }

  return (
    "<c:pieChart>" +
    '<c:varyColors val="1"/>' +
    '<c:ser><c:idx val="0"/><c:order val="0"/>' +
    seriesName(series, 0) +
    dPts +
    dataLabels(series, "bestFit", numberFormat) +
    categoryRef(spec) +
    valueRef(spec, series, 0, numberFormat) +
    "</c:ser>" +
    '<c:firstSliceAng val="0"/>' +
    "</c:pieChart>"
  );
}

function axes(spec: ChartSpec, numberFormat: string | null, zeroBase: boolean): string {
  const catPos = spec.horizontal ? "l" : "b";
  const valPos = spec.horizontal ? "b" : "l";

  return (
    "<c:catAx>" +
    '<c:axId val="' + CAT_AX + '"/>' +
    '<c:scaling><c:orientation val="minMax"/></c:scaling>' +
    '<c:delete val="0"/>' +
    '<c:axPos val="' + catPos + '"/>' +
    axisTitle(spec.categoryTitle, catPos) +
    '<c:majorTickMark val="out"/>' +
    '<c:minorTickMark val="none"/>' +
    '<c:tickLblPos val="nextTo"/>' +
    '<c:crossAx val="' + VAL_AX + '"/>' +
    '<c:crosses val="autoZero"/>' +
    '<c:auto val="1"/>' +
    '<c:lblAlgn val="ctr"/>' +
    '<c:lblOffset val="100"/>' +
    '<c:noMultiLvlLbl val="0"/>' +
    "</c:catAx>" +
    "<c:valAx>" +
    '<c:axId val="' + VAL_AX + '"/>' +
    '<c:scaling><c:orientation val="minMax"/>' + (zeroBase ? '<c:min val="0"/>' : "") + "</c:scaling>" +
    '<c:delete val="0"/>' +
    '<c:axPos val="' + valPos + '"/>' +
    '<c:majorGridlines><c:spPr><a:ln w="9525">' + solid(GRID) + "</a:ln></c:spPr></c:majorGridlines>" +
    axisTitle(spec.valueTitle, valPos) +
    (numberFormat !== null ? '<c:numFmt formatCode="' + Xml.attr(numberFormat) + '" sourceLinked="0"/>' : "") +
    '<c:majorTickMark val="none"/>' +
    '<c:minorTickMark val="none"/>' +
    '<c:tickLblPos val="nextTo"/>' +
    '<c:crossAx val="' + CAT_AX + '"/>' +
    '<c:crosses val="autoZero"/>' +
    '<c:crossBetween val="between"/>' +
    "</c:valAx>"
  );
}

// ── Fragments ───────────────────────────────────────────────────────────────

function rich(text: string, vertical: boolean): string {
  return (
    "<c:tx><c:rich>" +
    (vertical ? '<a:bodyPr rot="-5400000" vert="horz"/>' : "<a:bodyPr/>") +
    "<a:p><a:r><a:t>" + Xml.text(text) + "</a:t></a:r></a:p>" +
    "</c:rich></c:tx>"
  );
}

function axisTitle(text: string, position: string): string {
  return text === "" ? "" : "<c:title>" + rich(text, position === "l") + '<c:overlay val="0"/></c:title>';
}

function solid(hex: string): string {
  return '<a:solidFill><a:srgbClr val="' + hex + '"/></a:solidFill>';
}

function marker(color: string): string {
  return '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>' + solid(color) + "<a:ln>" + solid(color) + "</a:ln></c:spPr></c:marker>";
}

function seriesName(series: ChartSeriesSpec, idx: number): string {
  return (
    "<c:tx><c:strRef><c:f>Sheet1!$" + chartColumn(idx + 1) + "$1</c:f>" +
    '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + Xml.text(series.name) + "</c:v></c:pt></c:strCache>" +
    "</c:strRef></c:tx>"
  );
}

function dataLabels(series: ChartSeriesSpec, position: string, numberFormat: string | null): string {
  if (!series.labels) return "";
  return (
    "<c:dLbls>" +
    (numberFormat !== null ? '<c:numFmt formatCode="' + Xml.attr(numberFormat) + '" sourceLinked="0"/>' : "") +
    "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>" +
    '<c:dLblPos val="' + position + '"/>' +
    '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>' +
    "</c:dLbls>"
  );
}

function categoryRef(spec: ChartSpec): string {
  const rows = rowCount(spec);
  let pts = "";
  for (let r = 0; r < rows; r++) {
    pts += '<c:pt idx="' + r + '"><c:v>' + Xml.text(category(spec, r)) + "</c:v></c:pt>";
  }
  return (
    "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$" + (rows + 1) + "</c:f>" +
    '<c:strCache><c:ptCount val="' + rows + '"/>' + pts + "</c:strCache>" +
    "</c:strRef></c:cat>"
  );
}

function valueRef(spec: ChartSpec, series: ChartSeriesSpec, idx: number, numberFormat: string | null): string {
  const rows = rowCount(spec);
  const col = chartColumn(idx + 1);
  let pts = "";
  series.values.forEach((value, i) => {
    pts += '<c:pt idx="' + i + '"><c:v>' + chartNum(value) + "</c:v></c:pt>";
  });
  return (
    "<c:val><c:numRef><c:f>Sheet1!$" + col + "$2:$" + col + "$" + (rows + 1) + "</c:f>" +
    "<c:numCache><c:formatCode>" + Xml.text(numberFormat ?? "General") + "</c:formatCode>" +
    '<c:ptCount val="' + rows + '"/>' + pts + "</c:numCache>" +
    "</c:numRef></c:val>"
  );
}

/** Rows of data: the longer of the category list and the longest series. */
function rowCount(spec: ChartSpec): number {
  let rows = spec.categories.length;
  for (const series of spec.series) rows = Math.max(rows, series.values.length);
  return rows;
}

/** Row `r`'s category label; with no categories at all, the rows are numbered. */
function category(spec: ChartSpec, r: number): string {
  if (spec.categories.length === 0) return String(r + 1);
  return spec.categories[r] ?? "";
}

function stringCell(ref: string, text: string): string {
  return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + Xml.text(text) + "</t></is></c>";
}

/** Every plotted value is zero or more, so a bar's value axis starts at zero. */
function nonNegative(spec: ChartSpec): boolean {
  return spec.series.every((s) => s.values.every((v) => v >= 0));
}
