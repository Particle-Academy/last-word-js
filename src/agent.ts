/**
 * Agent — the structured-tool surface for LastWord. Mirrors PHP `Agent`.
 * Universal methods are synchronous; file-touching methods (`write`) are async
 * and Node-only (browsers have no sync FS).
 */

import { SchemaException, UnsupportedFormatException } from "./exceptions";
import { DocxTemplate } from "./writer/docx-template";
import { TemplateException } from "./exceptions";
import { fromMarkdown } from "./markdown/from-markdown";
import { toMarkdown } from "./markdown/to-markdown";
import { DocDiff } from "./ops/doc-diff";
import { DocOpSchema } from "./ops/doc-op-schema";
import { DocReducer } from "./ops/doc-reducer";
import { isList, valuesOf } from "./ops/php";
import type { DocOp } from "./ops/types";
import { DocReader } from "./reader/doc-reader";
import { DocxReader } from "./reader/docx-reader";
import { Format, detectFormat } from "./reader/format";
import { OdtReader } from "./reader/odt-reader";
import { RtfReader } from "./reader/rtf-reader";
import { Repairer } from "./schema/repairer";
import { Schema } from "./schema/schema";
import type { Block, Doc, ListItem, RepairResult, ValidationError, WriteResult } from "./schema/types";
import { Validator } from "./schema/validator";
import { DocxWriter } from "./writer/docx-writer";

/** This package's own version, pinned to package.json by `version.test.ts`. */
export const VERSION = "0.7.0";

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

/**
 * Write options. `template` is a `.dotx`/`.docx`: bytes anywhere, or a filesystem
 * path through `Agent.write()` (Node only — `toBytes()` is universal and a
 * browser has no filesystem).
 */
export type WriteOptions = {
  template?: Uint8Array | string;
};

/**
 * The `template` option as an opened template, or undefined when absent.
 *
 * An empty value counts as absent, so a host passing through a blank form field
 * gets the built-in look rather than an exception.
 */
function templateFrom(options?: WriteOptions): DocxTemplate | undefined {
  const template = options?.template;

  if (template === undefined || template === null) return undefined;

  if (typeof template === "string") {
    // Only `write()` resolves a path; reaching here with one means `toBytes()`
    // was given a path it cannot read.
    throw new TemplateException(
      "toBytes() needs the template as bytes — it is universal and cannot read a path. " +
        "Read the file yourself, or use write(), which accepts either.",
    );
  }

  if (template.length === 0) return undefined;

  return DocxTemplate.open(template);
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

  /**
   * DOCX bytes for a doc (no temp file). Universal. Throws SchemaException if invalid.
   *
   * `options.template` is a `.dotx`/`.docx` whose `word/styles.xml` and
   * `word/theme/theme1.xml` the document renders onto, so it comes out in a house
   * look rather than the built-in one (last-word#3).
   *
   * Binds BY STYLE NAME, with nothing to configure: the document model already
   * uses Word's own style ids, so a template defining `Normal`, `Heading1..n`,
   * `Quote` and `Hyperlink` binds on its own. Definitions the template lacks are
   * supplied from the built-in set, because a `w:pStyle` naming an undefined
   * style renders UNSTYLED in Word rather than erroring.
   *
   * NOT taken from the template: `w:sectPr` (page size, margins, headers,
   * footers), `word/numbering.xml` (this document's lists reference numbering ids
   * defined here) and `word/settings.xml`. So list markers and page setup stay
   * ours; the typography, colours and theme are the template's.
   *
   * An unusable template throws `TemplateException` rather than falling back — a
   * document that silently comes out in the wrong style is the failure this
   * option exists to end, and it lets a host validate a customer-supplied
   * template at upload rather than at render.
   *
   * Bytes rather than a path, because this entry point is universal and a browser
   * has no filesystem. `write()` accepts either.
   */
  toBytes(doc: Any, options?: WriteOptions): Uint8Array {
    assertValid(doc);
    return new DocxWriter(templateFrom(options)).toBytes(doc);
  },

  /**
   * Write a doc to disk as a .docx file (Node only). Throws SchemaException if
   * invalid. Same options as {@link toBytes}, except `template` may also be a
   * filesystem path.
   */
  async write(doc: Any, path: string, options?: WriteOptions): Promise<WriteResult> {
    assertValid(doc);
    const fs = await import("node:fs");

    let resolved = options;
    if (typeof options?.template === "string") {
      resolved = { ...options, template: fs.readFileSync(options.template) };
    }

    const bytes = new DocxWriter(templateFrom(resolved)).toBytes(doc);
    fs.writeFileSync(path, bytes);
    return { path, bytes: bytes.length, blocks: doc?.blocks?.length ?? 0 };
  },

  /**
   * Read a document back into the Doc model. Universal. The format is decided
   * from the CONTENT: .docx, legacy .doc (Word 97-2003), .odt and .rtf all
   * return the same shape. Mirrors PHP `Agent::read()`.
   *
   * Throws `UnsupportedFormatException` for bytes that are none of those,
   * naming what they are when that is knowable (`xls`, `pptx`, …), and a plain
   * `Error` for a file in a supported format that is damaged.
   */
  read(input: Uint8Array | ArrayBuffer): Doc {
    const bytes = toU8(input);
    const format = detectFormat(bytes);
    switch (format) {
      case Format.DOCX:
        return new DocxReader().read(bytes);
      case Format.ODT:
        return new OdtReader().read(bytes);
      case Format.RTF:
        return new RtfReader().read(bytes);
      case Format.DOC:
        // A compound file: DocReader reads a Word document and names anything
        // else (.xls, .ppt, .msg) itself, because only the container knows.
        return new DocReader().read(bytes);
      case Format.UNKNOWN:
        throw new UnsupportedFormatException(
          Format.UNKNOWN,
          "Agent.read() could not recognise these bytes as a document. It reads .docx, .doc (Word 97-2003), .odt and .rtf.",
        );
      default:
        throw new UnsupportedFormatException(
          format,
          `This is a .${format} file, not a word-processing document. Agent.read() reads .docx, .doc (Word 97-2003), .odt and .rtf.`,
        );
    }
  },

  /** Alias for {@see read}. */
  fromBytes(input: Uint8Array | ArrayBuffer): Doc {
    return this.read(input);
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

  /**
   * The ops that turn document `a` into document `b`. Mirrors PHP `Agent::diff`,
   * and gives the same ops in the same order for the same inputs.
   *
   * - `reduce(a, diff(a, b))` equals `b` (key order aside). The ops are verified
   *   by replaying them; ops that do not reproduce `b` become one `doc.replace`.
   * - Documents that write the same file diff to `[]`, so
   *   `diff(d, read(toBytes(d)))` is `[]`: a save without a change records
   *   nothing, even where the reader normalises (merged runs, a header row's
   *   bold, a dropped empty paragraph).
   * - Rewording one paragraph is one `blocks.replace` at its own path, even
   *   inside a table cell, a quote or a list; moving one is one `blocks.move`.
   *
   * Store `diff(newer, older)` to keep a version as the ops that restore it.
   * Both documents must be valid: the "same file" check writes them.
   */
  diff(a: Any, b: Any): DocOp[] {
    return DocDiff.diff(a, b);
  },

  /**
   * Apply one op, or a list of them, to a document; returns a new document and
   * never modifies the input. An op whose path or index does not resolve is
   * skipped.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  reduce(doc: Any, opOrOps: DocOp | readonly DocOp[]): Record<string, any> {
    // PHP's `$opOrOps === [] || array_is_list($opOrOps)`: an array, or an
    // object keyed exactly "0".."n-1", which is how PHP sees `{}` too.
    return DocReducer.applyAll(doc, isList(opOrOps) ? valuesOf(opOrOps) : [opOrOps]);
  },

  /** JSON Schema for one document op. */
  opSchema(): Record<string, unknown> {
    return DocOpSchema.jsonSchema();
  },

  /**
   * Whether two documents write the same file: runs the reader merges, a
   * header row's bold and an empty paragraph the writer drops do not make them
   * different.
   */
  equivalent(a: Any, b: Any): boolean {
    return DocDiff.equivalent(a, b);
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
