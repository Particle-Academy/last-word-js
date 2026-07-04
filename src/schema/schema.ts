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
  MAX_HEADING_LEVEL: 6,

  docRequiredKeys(): string[] {
    return ["blocks"];
  },

  jsonSchema(): Record<string, unknown> {
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
        link: { type: "string" },
        color: { type: "string", pattern: "^#[0-9a-fA-F]{6}$" },
        highlight: { type: "string", pattern: "^#[0-9a-fA-F]{6}$" },
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
      },
      definitions: {
        run: runSchema,
        listItem: {
          type: "object",
          required: ["runs"],
          properties: {
            runs: { type: "array", items: { $ref: "#/definitions/run" } },
            children: { type: "array", items: { $ref: "#/definitions/listItem" } },
          },
        },
        block: {
          type: "object",
          required: ["type"],
          properties: {
            type: { type: "string", enum: [...this.BLOCK_TYPES] },
            // heading
            level: { type: "integer", minimum: 1, maximum: this.MAX_HEADING_LEVEL },
            // heading + paragraph
            runs: { type: "array", items: { $ref: "#/definitions/run" } },
            align: { type: "string", enum: [...this.ALIGNMENTS] },
            // list
            ordered: { type: "boolean" },
            items: { type: "array", items: { $ref: "#/definitions/listItem" } },
            // table
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
