/**
 * Schema constants describing the Doc shape. Mirrors PHP `Schema\Schema`.
 */
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
              description: "Table width as a percentage of the text column.",
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
          },
        },
      },
    };
  },
};
