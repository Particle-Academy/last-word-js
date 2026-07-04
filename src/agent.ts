/**
 * Agent — the structured-tool surface for LastWord. Mirrors PHP `Agent`.
 * Universal methods are synchronous; file-touching methods (`write`) are async
 * and Node-only (browsers have no sync FS).
 */

import { SchemaException } from "./exceptions";
import { fromMarkdown } from "./markdown/from-markdown";
import { toMarkdown } from "./markdown/to-markdown";
import { DocxReader } from "./reader/docx-reader";
import { Repairer } from "./schema/repairer";
import { Schema } from "./schema/schema";
import type { Block, Doc, ListItem, RepairResult, ValidationError, WriteResult } from "./schema/types";
import { Validator } from "./schema/validator";
import { DocxWriter } from "./writer/docx-writer";

/** Feature-parity baseline with PHP last-word; bumped independently on npm. */
export const VERSION = "0.2.0";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function toU8(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input);
}

function assertValid(doc: Any): void {
  const errors = new Validator().validate(doc);
  if (errors.length > 0) {
    throw new SchemaException(
      "Doc failed schema validation. Call Agent.validateAndRepair() for a recoverable form.",
      errors,
    );
  }
}

export const Agent = {
  /** Validate a doc without writing. Empty array = valid. */
  validate(doc: Any): ValidationError[] {
    return new Validator().validate(doc);
  },

  /**
   * Validate + apply heuristic repairs (coerce strings to runs, clamp heading
   * levels, drop unknown block types with the error retained, default missing
   * blocks to []). Returns `{ ok, schema, errors }` where `ok` is true when
   * the repaired doc validates clean; `errors` retains anything the repair
   * dropped plus any remaining validation errors.
   */
  validateAndRepair(doc: Any): RepairResult {
    const errors = this.validate(doc);
    if (errors.length === 0) {
      return { ok: true, schema: doc, errors: [] };
    }
    const repairer = new Repairer();
    const repaired = repairer.repair(doc) as Doc;
    const remaining = this.validate(repaired);

    return {
      ok: remaining.length === 0,
      schema: repaired,
      errors: [...repairer.notes, ...remaining],
    };
  },

  /** DOCX bytes for a doc (no temp file). Universal. Throws SchemaException if invalid. */
  toBytes(doc: Any): Uint8Array {
    assertValid(doc);
    return new DocxWriter().toBytes(doc);
  },

  /** Write a doc to disk as a .docx file (Node only). Throws SchemaException if invalid. */
  async write(doc: Any, path: string): Promise<WriteResult> {
    assertValid(doc);
    const bytes = new DocxWriter().toBytes(doc);
    const fs = await import("node:fs");
    fs.writeFileSync(path, bytes);
    return { path, bytes: bytes.length, blocks: doc?.blocks?.length ?? 0 };
  },

  /** Read .docx bytes back into the Doc model. Universal. */
  read(input: Uint8Array | ArrayBuffer): Doc {
    return new DocxReader().read(toU8(input));
  },

  /** Alias for {@see read}. */
  fromBytes(input: Uint8Array | ArrayBuffer): Doc {
    return new DocxReader().read(toU8(input));
  },

  /** Doc → GFM markdown (the Editor bridge). */
  toMarkdown(doc: Any): string {
    return toMarkdown(doc);
  },

  /** GFM markdown → Doc (the Editor bridge). */
  fromMarkdown(markdown: string): Doc {
    return fromMarkdown(markdown);
  },

  /** Plain-text summary of a doc: title, block counts by type, word count. */
  describe(doc: Any): string {
    const title = String(doc?.title ?? "Untitled");
    const blocks: Any[] = Array.isArray(doc?.blocks) ? doc.blocks : [];

    const counts: Record<string, number> = {};
    for (const block of blocks) {
      const type = String(block?.type ?? "unknown");
      counts[type] = (counts[type] ?? 0) + 1;
    }

    const lines = [`Doc: ${title}`, `Blocks: ${blocks.length}`];
    const keys = Object.keys(counts);
    if (keys.length > 0) {
      lines.push("Kinds: " + keys.map((type) => `${counts[type]} ${type}`).join(", "));
    }
    lines.push(`Words: ${countWords(blocks)}`);

    return lines.join("\n");
  },

  /** JSON Schema export for LLM tool-use registration. */
  jsonSchema(): Record<string, unknown> {
    return Schema.jsonSchema();
  },

  version(): string {
    return VERSION;
  },
};

/** Prose word count: every run in headings, paragraphs, lists, tables, quotes. */
function countWords(blocks: Block[]): number {
  let text = "";
  const visitRuns = (runs: Any[] | undefined): void => {
    for (const run of runs ?? []) {
      if (typeof run?.text === "string") text += run.text + " ";
    }
  };
  const visitItems = (items: ListItem[] | undefined): void => {
    for (const item of items ?? []) {
      visitRuns(item.runs as Any[]);
      visitItems(item.children);
    }
  };
  const visit = (list: Any[]): void => {
    for (const block of list ?? []) {
      switch (block?.type) {
        case "heading":
        case "paragraph":
          visitRuns(block.runs);
          break;
        case "list":
          visitItems(block.items);
          break;
        case "table":
          for (const row of block.rows ?? []) {
            for (const cell of row?.cells ?? []) visit(cell?.blocks ?? []);
          }
          break;
        case "quote":
          visit(block.blocks ?? []);
          break;
        default:
          break;
      }
    }
  };
  visit(blocks as Any[]);
  return text.split(/\s+/).filter(Boolean).length;
}
