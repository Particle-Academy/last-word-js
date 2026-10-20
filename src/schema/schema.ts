/**
 * Schema constants describing the Doc shape. Mirrors PHP `Schema\Schema`.
 */
import { SUPPORTED_CHART_TYPES } from "../helpers/chart";

export const Schema = {
  VERSION: "0.1.0",

  BLOCK_TYPES: [
    "heading",
    "paragraph",
    "list",
    "table",
    "code",
    "quote",
    "image",
    "chart",
    "pageBreak",
    "hr",
  ] as const,
  ALIGNMENTS: ["left", "center", "right", "justify"] as const,
  BORDER_STYLES: ["single", "double", "dashed", "dotted", "none"] as const,
  PAGE_SIZES: ["letter", "legal", "a4"] as const,
  MAX_HEADING_LEVEL: 6,

  docRequiredKeys(): string[] {
    return ["blocks"];
  },

  jsonSchema(): Record<string, unknown> {
    const hex = { type: "string", pattern: "^#[0-9a-fA-F]{6}$" };
    const points = { type: "number", description: "Points." };
    const borderRef = { $ref: "#/definitions/border" };
    const fill = (keys: string[], value: unknown): Record<string, unknown> =>
      Object.fromEntries(keys.map((k) => [k, value]));

    const border = {
      type: "object",
      description: 'One border edge. {"style":"none"} REMOVES a border — a zero width is not it.',
      properties: {
        style: { type: "string", enum: [...this.BORDER_STYLES] },
        width: { type: "number", exclusiveMinimum: 0, description: "Points. Defaults to 0.5 (a hairline)." },
        color: hex,
      },
    };
    const boxBorders = {
      type: "object",
      description: "Box edges. Anything omitted is left alone rather than reset.",
      properties: fill(["top", "right", "bottom", "left"], borderRef),
    };
    const tableBorders = {
      type: "object",
      description: "Table edges — the box, plus the two inside directions.",
      properties: fill(["top", "right", "bottom", "left", "insideH", "insideV"], borderRef),
    };
    const boxSides = {
      type: "object",
      description: "Box spacing in points. Anything omitted is left alone.",
      properties: fill(["top", "right", "bottom", "left"], points),
    };
    const bandSlot = {
      type: "string",
      description:
        'Text. `{page}` is the current page number and `{pages}` the page count, both live fields: "Page {page} of {pages}".',
    };
    const band = {
      type: "object",
      additionalProperties: false,
      description:
        'A running header or footer: one line with a left, a centre and a right slot, e.g. {"left": "Q3 Board Memo | Confidential", "right": "Page {page} of {pages}"}. 9pt grey unless size/color say otherwise.',
      properties: {
        left: bandSlot,
        center: bandSlot,
        right: bandSlot,
        size: { type: "number", exclusiveMinimum: 0, description: "Points; 9 by default." },
        color: { type: "string", pattern: "^#[0-9A-Fa-f]{6}$", description: "Text colour; #6B7280 by default." },
      },
    };

    /** Properties every paragraph-shaped block accepts — paragraph, heading, list item. */
    const paragraphProps = {
      align: { type: "string", enum: [...this.ALIGNMENTS] },
      spaceBefore: points,
      spaceAfter: {
        type: "number",
        description: "Points. Zero is meaningful — the document default puts 8pt under every paragraph.",
      },
      lineHeight: { type: "number", exclusiveMinimum: 0, description: "A multiple of single spacing." },
      indentLeft: points,
      indentRight: points,
      keepNext: { type: "boolean", description: "Keep on the same page as the block after it." },
      shading: hex,
      borders: { $ref: "#/definitions/boxBorders" },
    };

    const runSchema = {
      type: "object",
      required: ["text"],
      properties: {
        text: { type: "string" },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        strike: { type: "boolean" },
        code: { type: "boolean" },
        smallCaps: { type: "boolean" },
        link: { type: "string" },
        color: hex,
        highlight: hex,
        size: {
          type: "number",
          exclusiveMinimum: 0,
          description: "Font size in points. Half-points are exactly representable.",
        },
        font: { type: "string", description: "Font family name." },
        letterSpacing: { type: "number", description: "Tracking in points; may be negative." },
      },
    };

    return {
      $schema: "http://json-schema.org/draft-07/schema#",
      title: "LastWord Doc",
      type: "object",
      required: this.docRequiredKeys(),
      properties: {
        title: { type: "string" },
        blocks: { type: "array", items: { $ref: "#/definitions/block" } },
        page: {
          type: "object",
          description:
            "Section geometry. A one-page business document does not fit inside the default one-inch margins.",
          properties: {
            size: { type: "string", enum: [...this.PAGE_SIZES] },
            orientation: { type: "string", enum: ["portrait", "landscape"] },
            margins: { $ref: "#/definitions/boxSides" },
            header: { $ref: "#/definitions/band", description: "A running header on every page." },
            footer: {
              $ref: "#/definitions/band",
              description: "A running footer on every page -- where a multi-page document carries its page numbers.",
            },
          },
        },
        defaultFont: { type: "string", description: "Font every run inherits unless it names its own." },
        defaultSize: {
          type: "number",
          exclusiveMinimum: 0,
          description: "Size in points every run inherits unless it names its own.",
        },
      },
      definitions: {
        run: runSchema,
        border,
        boxBorders,
        tableBorders,
        boxSides,
        band,
        listItem: {
          type: "object",
          required: ["runs"],
          properties: {
            runs: { type: "array", items: { $ref: "#/definitions/run" } },
            children: { type: "array", items: { $ref: "#/definitions/listItem" } },
            ...paragraphProps,
          },
        },
        block: {
          type: "object",
          required: ["type"],
          properties: {
            type: { type: "string", enum: [...this.BLOCK_TYPES] },
            // heading
            level: { type: "integer", minimum: 1, maximum: this.MAX_HEADING_LEVEL },
            // heading + paragraph (a heading is a paragraph and takes the
            // same properties, so a section label can be spaced and aligned
            // without being demoted to a bold paragraph)
            runs: { type: "array", items: { $ref: "#/definitions/run" } },
            ...paragraphProps,
            // list
            ordered: { type: "boolean" },
            items: { type: "array", items: { $ref: "#/definitions/listItem" } },
            // table
            widths: {
              type: "array",
              items: { type: "number", minimum: 0 },
              description:
                "Relative column weights — [30,40,30] and [3,4,3] are the same table. Also fixes the layout so Word honours them.",
            },
            width: {
              type: "number",
              exclusiveMinimum: 0,
              maximum: 100,
              description: "Table (or chart) width as a percentage of the text column.",
            },
            borders: { $ref: "#/definitions/tableBorders" },
            cellPadding: { $ref: "#/definitions/boxSides" },
            rows: {
              type: "array",
              items: {
                type: "object",
                required: ["cells"],
                properties: {
                  header: { type: "boolean" },
                  cells: {
                    type: "array",
                    items: {
                      type: "object",
                      required: ["blocks"],
                      properties: {
                        blocks: { type: "array", items: { $ref: "#/definitions/block" } },
                        shading: hex,
                        borders: { $ref: "#/definitions/boxBorders" },
                        padding: { $ref: "#/definitions/boxSides" },
                        valign: { type: "string", enum: ["top", "center", "bottom"] },
                        colSpan: { type: "integer", minimum: 1 },
                        rowSpan: {
                          type: "integer",
                          minimum: 1,
                          description:
                            "Written HTML-style: the cell appears ONCE, and the rows it covers list only their own remaining cells.",
                        },
                      },
                    },
                  },
                },
              },
            },
            // code
            language: { type: "string" },
            text: { type: "string" },
            // quote
            blocks: { type: "array", items: { $ref: "#/definitions/block" } },
            // image
            src: { type: "string" },
            widthPx: { type: "number", exclusiveMinimum: 0 },
            heightPx: { type: "number", exclusiveMinimum: 0 },
            alt: { type: "string" },
            // chart (also takes widthPx / heightPx / alt, and `width` as a percentage of the text column)
            option: chartOptionJsonSchema(),
            numberFormat: {
              type: "string",
              description:
                'A chart\'s Excel number format for the values, on the value axis, the data labels and the workbook: "#,##0", "0.0%", "$#,##0.00". Absent, General.',
            },
          },
        },
      },
    };
  },
};

/**
 * A chart block's `option`: the subset of an ECharts option this writer reads,
 * in dark-slide's words wherever the two agree. Mirrors PHP
 * `Schema::chartOptionJsonSchema()`.
 */
function chartOptionJsonSchema(): Record<string, unknown> {
  const hex = { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" };
  const axis = {
    type: "object",
    properties: {
      type: { enum: ["category", "value"] },
      data: { type: "array", items: { type: ["string", "number"] }, description: "The category labels, on the category axis." },
      name: { type: "string", description: "The axis title." },
    },
  };
  return {
    type: "object",
    required: ["series"],
    description:
      'A chart block\'s Apache ECharts option -- the same `option` a dark-slide chart element takes. A chart is a native Word chart, editable in Word with its numbers in an embedded workbook. Column chart: {"xAxis": {"type": "category", "data": ["Q1", "Q2"]}, "yAxis": {"type": "value"}, "series": [{"type": "bar", "name": "Revenue", "data": [120, 135]}]}. Swap the two axes for horizontal bars. Pie: {"series": [{"type": "pie", "data": [{"name": "Payroll", "value": 61}, {"name": "Cloud", "value": 22}]}]}.',
    properties: {
      series: {
        type: "array",
        minItems: 1,
        description: "The data to plot, every series of one type; a pie takes exactly one.",
        items: {
          type: "object",
          required: ["data"],
          properties: {
            type: {
              type: "string",
              enum: [...SUPPORTED_CHART_TYPES],
              description:
                "The chart kind, bar when absent. A bar series draws columns over a category xAxis and horizontal bars over a category yAxis. scatter, radar, gauge and other ECharts types are not supported in a document.",
            },
            name: { type: "string", description: "The series name, shown in the legend and the workbook header." },
            data: {
              type: "array",
              description:
                'The points, one per category: a number, or {"value": number}; for a pie, {"name": string, "value": number}, whose names become the labels. Colour ONE point to make the finding stand out: {"value": 90.7, "itemStyle": {"color": "#DC2626"}}.',
            },
            itemStyle: {
              type: "object",
              properties: { color: hex },
              description: 'The series\' own colour: {"color": "#1F3A5F"}. A muted series with one highlighted point reads faster than a rainbow.',
            },
            label: {
              type: "object",
              properties: { show: { type: "boolean" } },
              description: 'Data labels: {"show": true} draws each value on the chart, in numberFormat.',
            },
            smooth: { type: "boolean", description: "A smoothed line (line series)." },
            stack: { type: "string", description: "Any name stacks the bar series that carry one." },
          },
        },
      },
      xAxis: axis,
      yAxis: axis,
      title: {
        type: "object",
        properties: { text: { type: "string" } },
        description: 'The chart title, {"text": "..."}. A report usually captions the chart in a paragraph instead.',
      },
      legend: {
        type: "object",
        properties: { show: { type: "boolean" } },
        description: "A legend below the chart, drawn when present and show is not false. A single series rarely needs one.",
      },
      color: { type: "array", items: hex, description: "The palette series (or pie slices) take in turn. Absent, Office's default accents." },
    },
  };
}
