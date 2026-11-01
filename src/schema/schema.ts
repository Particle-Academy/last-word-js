/**
 * Schema constants describing the Doc shape. Mirrors PHP `Schema\Schema`.
 */
import { SUPPORTED_CHART_TYPES } from "../helpers/chart";
import type { FieldToken, ResolvedPageMargins } from "./types";

const isFiniteNumber = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);

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

  /**
   * Block types a footnote's `blocks` refuse, at any depth: a page break means
   * nothing at the foot of a page, and a chart's parts are wired to the body.
   * (`portable` is listed for parity with PHP; this engine has no portable
   * block, so it is refused as an unknown type in any case.) A footnote inside
   * a footnote is refused too (Word has no such thing), and so is a footnote in
   * a header or footer, which Word cannot hold.
   */
  FOOTNOTE_REFUSED_TYPES: ["chart", "pageBreak", "portable"] as const,

  /**
   * What a page margin is when the document does not say, in points: one inch
   * each side, and half an inch from the page edge to the header and to the
   * footer. The reader OMITS a value equal to these (a page read back carries
   * only what differs), so a viewer resolves the rest with `pageMargins()`.
   */
  PAGE_MARGIN_DEFAULT: 72,
  PAGE_BAND_DISTANCE_DEFAULT: 36,

  /**
   * The page's running bands, in the order the schema lists them and read()
   * returns them (PHP 0.16.0). `header` / `footer` are every page; with a
   * `firstHeader` they are every page but the first, and with an `evenHeader`
   * the odd pages. Each variant is the same shape as `header`.
   */
  PAGE_BANDS: ["header", "footer", "firstHeader", "firstFooter", "evenHeader", "evenFooter"] as const,

  /** A band key => [kind, Word reference type]. */
  PAGE_BAND_SLOTS: {
    header: ["header", "default"],
    footer: ["footer", "default"],
    firstHeader: ["header", "first"],
    firstFooter: ["footer", "first"],
    evenHeader: ["header", "even"],
    evenFooter: ["footer", "even"],
  } as const,

  /**
   * The field tokens live in header and footer text (PHP 0.16.0), as one regex:
   * `{page}` `{pages}` `{section}` `{sectionpages}` `{author}` `{title}`
   * `{filename}`, and `{date}` / `{time}` with an optional `:FORMAT` -- a Word
   * date-time picture, one or more characters, none of `{`, `}`, `"` or a
   * backslash. Names are lower-case and exact. Anything else in braces is
   * literal text, and so is `{page:x}`: only date and time take a format.
   * Group 1 is a plain name; group 2 a dated name and group 3 its format.
   * Mirrors PHP `Schema::FIELD_TOKEN`. Not global: use `splitFieldTokens()`.
   */
  FIELD_TOKEN: /\{(?:(page|pages|section|sectionpages|author|title|filename)|(date|time)(?::([^{}"\\]+))?)\}/,

  /** Field token name => Word field instruction. */
  FIELD_INSTRUCTIONS: {
    page: "PAGE",
    pages: "NUMPAGES",
    section: "SECTION",
    sectionpages: "SECTIONPAGES",
    author: "AUTHOR",
    title: "TITLE",
    filename: "FILENAME",
    date: "DATE",
    time: "TIME",
  } as Readonly<Record<string, string>>,

  /**
   * Header/footer text split into literal text and field tokens, in order: a
   * string is literal, an object is a field -- `{token, name, instruction}`,
   * where `instruction` is the Word field instruction without its surrounding
   * spaces (`PAGE`, `DATE \@ "yyyy-MM-dd"`). Empty pieces are dropped. Mirrors
   * PHP `Schema::splitFieldTokens()`.
   */
  splitFieldTokens(text: string): (string | FieldToken)[] {
    const out: (string | FieldToken)[] = [];
    let at = 0;
    for (const m of text.matchAll(new RegExp(this.FIELD_TOKEN.source, "g"))) {
      const token = m[0];
      const offset = m.index ?? 0;
      if (offset > at) out.push(text.slice(at, offset));
      const plain = m[1] ?? "";
      const name = plain !== "" ? plain : (m[2] ?? "");
      const format = plain === "" ? (m[3] ?? "") : "";
      out.push({
        token,
        name,
        instruction: this.FIELD_INSTRUCTIONS[name] + (format !== "" ? ` \\@ "${format}"` : ""),
      });
      at = offset + token.length;
    }
    if (at < text.length) out.push(text.slice(at));
    return out;
  },

  /**
   * A Word field instruction back as its token, or null when it is not one:
   * the first word names the field (any case), `\@ "FORMAT"` (or an unquoted
   * `\@ FORMAT`) is DATE's and TIME's picture, and every other switch or
   * argument is dropped (`\* MERGEFORMAT`, FILENAME's `\p`, AUTHOR's new name).
   * A picture the grammar cannot hold reads as the bare `{date}` / `{time}`.
   * Mirrors PHP `Schema::fieldTokenOf()`; whitespace is PCRE's `\s` (ASCII),
   * not JavaScript's wider one.
   */
  fieldTokenOf(instruction: string): string | null {
    const m = /^[ \t\n\v\f\r]*([A-Za-z]+)([\s\S]*)$/.exec(instruction);
    if (m === null) return null;
    const word = m[1]!.toUpperCase();
    const name = Object.keys(this.FIELD_INSTRUCTIONS).find((k) => this.FIELD_INSTRUCTIONS[k] === word);
    if (name === undefined) return null;
    if (name === "date" || name === "time") {
      const p = /\\@[ \t\n\v\f\r]*(?:"([^"]*)"|([^ \t\n\v\f\r]+))/.exec(m[2]!);
      if (p !== null) {
        const format = (p[1] ?? "") !== "" ? p[1]! : (p[2] ?? "");
        const token = `{${name}:${format}}`;
        const check = this.FIELD_TOKEN.exec(token);
        if (format !== "" && check !== null && check[0] === token) return token;
      }
    }
    return `{${name}}`;
  },

  /**
   * Every page margin and both header/footer distances, in points, defaults
   * filled in -- what a paged viewer lays a page out from. `read()` omits a
   * value equal to its default (`PAGE_MARGIN_DEFAULT`,
   * `PAGE_BAND_DISTANCE_DEFAULT`), and omits `page` altogether when nothing
   * differs, so this is the one place the defaults are applied. Mirrors PHP
   * `Schema::pageMargins()`.
   */
  pageMargins(doc: unknown): ResolvedPageMargins {
    const d = doc as { page?: unknown } | null;
    const page = d !== null && typeof d === "object" && !Array.isArray(d) && d.page !== null && typeof d.page === "object" && !Array.isArray(d.page)
      ? (d.page as Record<string, unknown>)
      : {};
    const m = page.margins !== null && typeof page.margins === "object" && !Array.isArray(page.margins)
      ? (page.margins as Record<string, unknown>)
      : {};
    const value = (k: string, fallback: number): number => (isFiniteNumber(m[k]) ? (m[k] as number) : fallback);

    return {
      top: value("top", this.PAGE_MARGIN_DEFAULT),
      right: value("right", this.PAGE_MARGIN_DEFAULT),
      bottom: value("bottom", this.PAGE_MARGIN_DEFAULT),
      left: value("left", this.PAGE_MARGIN_DEFAULT),
      header: value("header", this.PAGE_BAND_DISTANCE_DEFAULT),
      footer: value("footer", this.PAGE_BAND_DISTANCE_DEFAULT),
    };
  },

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
        'Text with live Word fields: {page} and {pages} ("Page {page} of {pages}"), {section}, {sectionpages}, {author} (meta.author), {title} (the document title), {filename}, {date} and {time} in the reader\'s locale, or {date:FORMAT} / {time:FORMAT} with a Word picture such as {date:d MMMM yyyy} or {time:HH:mm}. Anything else in braces is literal text.',
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
        footnote: { $ref: "#/definitions/footnote" },
      },
    };

    return {
      $schema: "http://json-schema.org/draft-07/schema#",
      title: "LastWord Doc",
      type: "object",
      required: this.docRequiredKeys(),
      properties: {
        title: {
          type: "string",
          description:
            "The document title: docProps/core.xml dc:title, what Word shows in File > Info and what a {title} field in a header or footer shows. Not rendered in the body; write a heading for that.",
        },
        meta: {
          type: "object",
          additionalProperties: false,
          description: "Document properties beyond the title (which is the top-level `title`).",
          properties: {
            author: {
              type: "string",
              description: "docProps/core.xml dc:creator: Word's Author property, and what an {author} field in a header or footer shows.",
            },
          },
        },
        blocks: { type: "array", items: { $ref: "#/definitions/block" } },
        page: {
          type: "object",
          description:
            "Section geometry. A one-page business document does not fit inside the default one-inch margins.",
          properties: {
            size: { type: "string", enum: [...this.PAGE_SIZES] },
            orientation: { type: "string", enum: ["portrait", "landscape"] },
            margins: { $ref: "#/definitions/pageMargins" },
            header: {
              $ref: "#/definitions/band",
              description: "A running header on every page. With a firstHeader, every page but the first; with an evenHeader, the odd pages.",
            },
            footer: {
              $ref: "#/definitions/band",
              description:
                "A running footer on every page -- where a multi-page document carries its page numbers. With a firstFooter, every page but the first; with an evenFooter, the odd pages.",
            },
            firstHeader: {
              oneOf: [{ $ref: "#/definitions/band" }, { $ref: "#/definitions/blankBand" }],
              description:
                'The header of the FIRST page only, in place of header there (Word\'s "Different first page"): a title-page header, a plainer one after. {"blocks": []} leaves page one with no header when header is set.',
            },
            firstFooter: {
              oneOf: [{ $ref: "#/definitions/band" }, { $ref: "#/definitions/blankBand" }],
              description:
                'The footer of the first page only, in place of footer there. Without it the footer is on the first page too. {"blocks": []} leaves page one with no footer when footer is set.',
            },
            evenHeader: {
              oneOf: [{ $ref: "#/definitions/band" }, { $ref: "#/definitions/blankBand" }],
              description:
                'The header of even pages (Word\'s "Different odd and even pages"); header is then the odd pages\'. For a bound, two-sided document.',
            },
            evenFooter: {
              oneOf: [{ $ref: "#/definitions/band" }, { $ref: "#/definitions/blankBand" }],
              description: "The footer of even pages; footer is then the odd pages'. Without it the footer is on even pages too.",
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
        pageMargins: {
          type: "object",
          additionalProperties: false,
          description:
            "Page margins in points, 72 (one inch) each side by default, and the header and footer DISTANCES: how far the header sits from the top edge of the page and the footer from the bottom edge, 36 by default. Word moves the body down (or up) when a header (or footer) is taller than the margin leaves room for. Anything omitted is the default.",
          properties: {
            top: points,
            right: points,
            bottom: points,
            left: points,
            header: { type: "number", minimum: 0, description: "Points from the top edge of the page to the top of the header. 36 by default." },
            footer: { type: "number", minimum: 0, description: "Points from the bottom edge of the page to the bottom of the footer. 36 by default." },
          },
        },
        footnote: {
          type: "object",
          required: ["blocks"],
          additionalProperties: false,
          description:
            'A footnote, carried by a run whose text is "" and which sits where the reference mark goes: {"text": "", "footnote": {"blocks": [{"type": "paragraph", "runs": [{"text": "Audited figures, FY2026."}]}]}}. Word numbers the marks 1, 2, 3 in document order and sets each footnote at the foot of the page that holds its mark. The blocks are the body vocabulary less chart and pageBreak, with no footnote inside a footnote. Not in a header or footer.',
          properties: {
            blocks: {
              type: "array",
              items: {
                allOf: [
                  { $ref: "#/definitions/block" },
                  { not: { properties: { type: { enum: ["chart", "pageBreak"] } }, required: ["type"] } },
                ],
              },
            },
          },
        },
        band,
        blankBand: {
          type: "object",
          required: ["blocks"],
          additionalProperties: false,
          description:
            'An empty first or even band, {"blocks": []}: that page shows no header (footer) where the default would otherwise show. (This engine writes no {blocks} band with content.)',
          properties: { blocks: { type: "array", maxItems: 0 } },
        },
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
