import { clone, isNumeric, isPlainObject } from "../util";
import { Schema } from "./schema";
import type { ValidationError } from "./types";
import { hasFootnote, validateChart } from "./validator";
import { chartLabel } from "../helpers/chart";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** A scalar as PHP's `(string)` cast gives it (true "1", false ""); null for anything else. */
function scalarText(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : null;
  if (typeof value === "boolean") return value ? "1" : "";
  return null;
}

/** `{chartType, title, categories, series: [{name, values, color}]}` as the chart block it stands for. */
function chartFromShorthand(block: Any): Record<string, unknown> {
  const chartType = typeof block.chartType === "string" ? block.chartType.toLowerCase() : "column";
  const type = chartType === "line" ? "line" : chartType === "pie" ? "pie" : "bar";
  const categories: Any[] = Array.isArray(block.categories) ? block.categories : [];

  const series: Record<string, unknown>[] = [];
  for (const one of block.series as Any[]) {
    if (!isPlainObject(one)) continue;
    const o = one as Any;
    const out: Record<string, unknown> = { type };
    if (o.name !== undefined) out.name = o.name;
    const values: Any[] = Array.isArray(o.values) ? o.values : Array.isArray(o.data) ? o.data : [];
    out.data = type === "pie" ? values.map((value, i) => ({ name: chartLabel(categories[i]) ?? `Slice ${i + 1}`, value })) : values;
    if (typeof o.color === "string") out.itemStyle = { color: o.color };
    series.push(out);
  }

  const option: Record<string, unknown> = {};
  if (typeof block.title === "string" && block.title !== "") option.title = { text: block.title };
  if (type !== "pie") {
    const category = { type: "category", data: categories };
    const value = { type: "value" };
    if (chartType === "bar") {
      option.xAxis = value;
      option.yAxis = category;
    } else {
      option.xAxis = category;
      option.yAxis = value;
    }
  }
  option.series = series;

  const out: Record<string, unknown> = { type: "chart", option };
  for (const key of ["widthPx", "heightPx", "width", "numberFormat", "alt"]) {
    if (key in block) out[key] = block[key];
  }
  return out;
}

/**
 * Heuristic doc repair. Mirrors PHP `Schema\Repairer`. Never mutates input.
 *
 * Heuristics:
 *  - missing / non-array `blocks` → `[]`
 *  - bare strings in `blocks` → paragraphs; `runs: "text"` → `[{ text }]`;
 *    string entries inside a runs array → `{ text }`
 *  - heading levels clamped to 1..6 (non-numeric → 1)
 *  - unknown block types dropped, with the error retained in `notes`
 *  - list items given as strings → `{ runs: [{ text }] }`
 */
export class Repairer {
  /** Errors describing content the last `repair()` call had to drop. */
  notes: ValidationError[] = [];

  /** 'body', 'band' or 'footnote': where a footnote run is being repaired, and so whether it may stay. */
  private context: "body" | "band" | "footnote" = "body";

  repair(docInput: Any): Record<string, unknown> {
    this.notes = [];
    this.context = "body";
    const doc: Any = isPlainObject(docInput) ? clone(docInput) : {};

    if (doc.title !== undefined && typeof doc.title !== "string") {
      doc.title = String(doc.title);
    }
    let out: Any = doc;
    if ("meta" in doc) out = this.repairMeta(doc);
    out.blocks = this.repairBlocks(out.blocks ?? [], "/blocks");

    // A `{blocks}` header or footer (PHP 0.14.0; this engine does not write one)
    // is repaired in the band's context, so a footnote in it is dropped with a
    // note, as PHP does: a header or footer cannot hold one. Every band key,
    // first and even included (PHP 0.16.0).
    if (isPlainObject(out.page)) {
      for (const kind of Schema.PAGE_BANDS) {
        const band = out.page[kind];
        if (isPlainObject(band) && Array.isArray((band as Any).blocks)) {
          this.context = "band";
          try {
            (band as Any).blocks = this.repairBlocks((band as Any).blocks, `/page/${kind}/blocks`);
          } finally {
            this.context = "body";
          }
        }
      }
    }

    return out;
  }

  /**
   * `meta` (PHP 0.16.0): `meta.title` moves to the top-level `title` when there
   * is none (and is dropped when there is: the top-level one is the title); a
   * non-text author is coerced when it is a scalar and dropped otherwise; a
   * `meta` that is not an object is dropped, each with a note. Mirrors PHP
   * `Repairer::repairMeta()`, the title placed first as PHP places it.
   */
  private repairMeta(doc: Any): Any {
    const meta = doc.meta;
    if (Array.isArray(meta) && meta.length === 0) {
      doc.meta = {}; // PHP's empty object
      return doc;
    }
    if (!isPlainObject(meta)) {
      delete doc.meta;
      this.notes.push({ path: "/meta", message: "Dropped a meta that is not an object." });
      return doc;
    }
    const m: Any = { ...meta };
    let out: Any = doc;
    if ("title" in m) {
      if (!("title" in doc) && typeof m.title === "string") {
        out = { title: m.title, ...doc };
        this.notes.push({ path: "/meta/title", message: "Moved meta.title to the document title." });
      } else {
        this.notes.push({ path: "/meta/title", message: 'Dropped meta.title: the document title is the top-level "title".' });
      }
      delete m.title;
    }
    if ("author" in m && typeof m.author !== "string") {
      if (["number", "boolean", "bigint"].includes(typeof m.author)) {
        // PHP's (string) cast: true is "1", false is "".
        m.author = typeof m.author === "boolean" ? (m.author ? "1" : "") : String(m.author);
        this.notes.push({ path: "/meta/author", message: "Coerced a non-text author to text." });
      } else {
        delete m.author;
        this.notes.push({ path: "/meta/author", message: "Dropped a non-text author." });
      }
    }
    out.meta = m;
    return out;
  }

  private repairBlocks(blocks: Any, path: string): Record<string, unknown>[] {
    if (!Array.isArray(blocks)) return [];
    const out: Record<string, unknown>[] = [];
    blocks.forEach((block: Any, i: number) => {
      const repaired = this.repairBlock(block, `${path}/${i}`);
      if (repaired !== null) out.push(repaired);
    });
    return out;
  }

  private repairBlock(block: Any, path: string): Record<string, unknown> | null {
    // A bare string block coerces to a paragraph.
    if (typeof block === "string") {
      return { type: "paragraph", runs: [{ text: block }] };
    }
    if (!isPlainObject(block)) {
      this.notes.push({ path, message: "Dropped non-object block during repair." });
      return null;
    }
    if (typeof block.type !== "string" || !(Schema.BLOCK_TYPES as readonly string[]).includes(block.type)) {
      // Common LLM slip: `{ type: "text", text: "…" }` → paragraph.
      if (typeof block.text === "string") {
        return { type: "paragraph", runs: [{ text: block.text }] };
      }
      this.notes.push({
        path: `${path}/type`,
        message: `Dropped block with unknown type \`${String(block.type)}\` during repair.`,
      });
      return null;
    }

    switch (block.type) {
      case "heading": {
        let level = isNumeric(block.level) ? Math.trunc(Number(block.level)) : 1;
        if (level < 1) level = 1;
        if (level > Schema.MAX_HEADING_LEVEL) level = Schema.MAX_HEADING_LEVEL;
        block.level = level;
        block.runs = this.repairRuns(block.runs, `${path}/runs`);
        break;
      }
      case "paragraph":
        block.runs = this.repairRuns(block.runs, `${path}/runs`);
        if (block.align !== undefined && !(Schema.ALIGNMENTS as readonly string[]).includes(block.align)) {
          delete block.align;
        }
        break;

      case "list":
        if (block.ordered !== undefined) block.ordered = Boolean(block.ordered);
        block.items = this.repairListItems(block.items, `${path}/items`);
        break;

      case "table":
        block.rows = this.repairRows(block.rows, path);
        break;

      case "code":
        block.text = typeof block.text === "string" ? block.text : String(block.text ?? "");
        if (block.language !== undefined && typeof block.language !== "string") delete block.language;
        break;

      case "quote":
        block.blocks = this.repairBlocks(block.blocks ?? [], `${path}/blocks`);
        break;

      case "image":
        if (typeof block.src !== "string" || !block.src.startsWith("data:image/")) {
          this.notes.push({ path: `${path}/src`, message: "Dropped image without a usable data URL during repair." });
          return null;
        }
        for (const key of ["widthPx", "heightPx"] as const) {
          if (block[key] !== undefined && (!isNumeric(block[key]) || Number(block[key]) <= 0)) {
            delete block[key];
          } else if (block[key] !== undefined) {
            block[key] = Number(block[key]);
          }
        }
        if (block.alt !== undefined && typeof block.alt !== "string") delete block.alt;
        break;

      case "chart":
        return this.repairChart(block, path);

      // A break keeps its keys, `page` and the near-miss ones above all: stripped,
      // any unrelated repair elsewhere would turn a refused landscape section into
      // `ok: true` and an all-portrait document.
      case "pageBreak":
      case "hr":
        break;
    }

    return block;
  }

  /**
   * A chart block. Repairs the SHORTHAND an agent reaches for --
   * `{chartType, title, categories, series: [{name, values, color}]}`, where
   * `chartType` takes Word's words (column, bar = horizontal, line, pie) -- into
   * the ECharts option it stands for, and `series[].values` into `data`. A chart
   * that still cannot be drawn is dropped, with the reason kept. Mirrors PHP
   * `Repairer::repairChart()`.
   */
  private repairChart(input: Any, path: string): Record<string, unknown> | null {
    let block: Any = input;
    if (!isPlainObject(block.option) && Array.isArray(block.series)) {
      block = chartFromShorthand(block);
      this.notes.push({ path: `${path}/option`, message: "Converted the chart shorthand (chartType / categories / series[].values) to an ECharts option." });
    }

    const option = block.option;
    if (isPlainObject(option) && Array.isArray((option as Any).series)) {
      (option as Any).series = (option as Any).series.map((series: Any, i: number) => {
        if (isPlainObject(series) && (series as Any).data === undefined && Array.isArray((series as Any).values)) {
          const { values, ...rest } = series as Any;
          this.notes.push({ path: `${path}/option/series/${i}/data`, message: 'Renamed series "values" to "data", as ECharts names it.' });
          return { ...rest, data: values };
        }
        return series;
      });
    }

    for (const key of ["widthPx", "heightPx"] as const) {
      if (block[key] !== undefined && block[key] !== null && !(typeof block[key] === "number" && Number.isFinite(block[key]) && block[key] > 0)) {
        delete block[key];
      }
    }
    if (block.width !== undefined && block.width !== null && !(typeof block.width === "number" && block.width > 0 && block.width <= 100)) {
      delete block.width;
    }
    if (block.numberFormat !== undefined && block.numberFormat !== null && !(typeof block.numberFormat === "string" && block.numberFormat !== "")) {
      delete block.numberFormat;
    }
    if (block.alt !== undefined && block.alt !== null && typeof block.alt !== "string") delete block.alt;

    if (validateChart(block, path).length > 0) {
      this.notes.push({ path, message: "Dropped a chart that cannot be drawn (see Agent.validate() for why)." });
      return null;
    }
    return block;
  }

  private repairRuns(runs: Any, path: string): Record<string, unknown>[] {
    // A bare string coerces to a single run.
    if (typeof runs === "string") return [{ text: runs }];
    if (!Array.isArray(runs)) return [];
    const out: Record<string, unknown>[] = [];
    for (const [i, run] of runs.entries()) {
      if (typeof run === "string") {
        out.push({ text: run });
        continue;
      }
      if (!isPlainObject(run)) continue;
      if (hasFootnote(run)) {
        out.push(...this.repairFootnoteRun(run, `${path}/${i}`));
        continue;
      }
      if (typeof run.text !== "string") run.text = String(run.text ?? "");
      for (const flag of ["bold", "italic", "underline", "strike", "code"] as const) {
        if (run[flag] !== undefined && typeof run[flag] !== "boolean") run[flag] = Boolean(run[flag]);
      }
      if (run.link !== undefined && typeof run.link !== "string") delete run.link;
      for (const key of ["color", "highlight"] as const) {
        if (run[key] !== undefined && (typeof run[key] !== "string" || !/^#[0-9a-fA-F]{6}$/.test(run[key]))) {
          delete run[key];
        }
      }
      out.push(run);
    }
    return out;
  }

  /**
   * A footnote run, repaired into what it should have been: text the agent
   * typed on the mark moves to a run of its own before it (losing nothing), a
   * missing `text` becomes "", a bare string or a bare list of blocks is
   * wrapped, and a footnote where Word cannot hold one (a header or footer,
   * another footnote) is dropped with a note. Mirrors PHP
   * `Repairer::repairFootnoteRun()`.
   */
  private repairFootnoteRun(input: Any, path: string): Record<string, unknown>[] {
    const { footnote, ...run } = input as Any;
    let note: Any = footnote;
    const out: Record<string, unknown>[] = [];
    const typed = scalarText(run.text);
    if (typed !== null && typed !== "") {
      run.text = typed;
      out.push(run);
      this.notes.push({ path: `${path}/text`, message: "Moved text typed on a footnote mark into a run of its own before it." });
    }

    if (this.context !== "body") {
      this.notes.push({
        path: `${path}/footnote`,
        message: this.context === "band" ? "Dropped a footnote: a header or footer cannot hold one." : "Dropped a footnote inside a footnote.",
      });
      return out;
    }

    if (typeof note === "string" && note !== "") {
      note = { blocks: [{ type: "paragraph", runs: [{ text: note }] }] };
      this.notes.push({ path: `${path}/footnote`, message: "Coerced a string footnote to {blocks: [paragraph]}." });
    } else if (Array.isArray(note) && note.length > 0) {
      note = { blocks: note };
      this.notes.push({ path: `${path}/footnote`, message: "Coerced a list of blocks to {blocks}." });
    }
    let blocks: Any = isPlainObject(note) ? (note as Any).blocks : null;
    if (typeof blocks === "string" && blocks !== "") {
      blocks = [{ type: "paragraph", runs: [{ text: blocks }] }];
    }
    if (!Array.isArray(blocks)) {
      this.notes.push({ path: `${path}/footnote`, message: "Dropped a footnote without blocks." });
      return out;
    }

    const at = `${path}/footnote/blocks`;
    this.context = "footnote";
    try {
      blocks = this.dropRefusedInFootnote(this.repairBlocks(blocks, at), at);
    } finally {
      this.context = "body";
    }
    if (blocks.length === 0) {
      this.notes.push({ path: `${path}/footnote`, message: "Dropped a footnote with no blocks left." });
      return out;
    }
    if (!Object.prototype.hasOwnProperty.call(run, "text")) {
      this.notes.push({ path: `${path}/text`, message: 'Defaulted a footnote run\'s text to "".' });
    }
    out.push({ text: "", footnote: { blocks } });
    return out;
  }

  /** A footnote's blocks without what a footnote cannot hold (`Schema.FOOTNOTE_REFUSED_TYPES`), at any depth, each drop noted. */
  private dropRefusedInFootnote(blocks: Record<string, unknown>[], path: string): Record<string, unknown>[] {
    const out: Record<string, unknown>[] = [];
    blocks.forEach((block: Any, i: number) => {
      const at = `${path}/${i}`;
      const type = block?.type;
      if ((Schema.FOOTNOTE_REFUSED_TYPES as readonly string[]).includes(type)) {
        this.notes.push({ path: at, message: `Dropped a ${type} block: a footnote cannot hold one.` });
        return;
      }
      if (type === "quote" && Array.isArray(block.blocks)) {
        block.blocks = this.dropRefusedInFootnote(block.blocks, `${at}/blocks`);
      }
      if (type === "table" && Array.isArray(block.rows)) {
        block.rows.forEach((row: Any, r: number) => {
          (Array.isArray(row?.cells) ? row.cells : []).forEach((cell: Any, c: number) => {
            if (Array.isArray(cell?.blocks)) {
              cell.blocks = this.dropRefusedInFootnote(cell.blocks, `${at}/rows/${r}/cells/${c}/blocks`);
            }
          });
        });
      }
      out.push(block);
    });
    return out;
  }

  private repairListItems(items: Any, path: string): Record<string, unknown>[] {
    if (!Array.isArray(items)) return [];
    const out: Record<string, unknown>[] = [];
    for (const [i, item] of items.entries()) {
      if (typeof item === "string") {
        out.push({ runs: [{ text: item }] });
        continue;
      }
      if (!isPlainObject(item)) continue;
      item.runs = this.repairRuns(item.runs, `${path}/${i}/runs`);
      if (item.children !== undefined) {
        const children = this.repairListItems(item.children, `${path}/${i}/children`);
        if (children.length > 0) item.children = children;
        else delete item.children;
      }
      out.push(item);
    }
    return out;
  }

  private repairRows(rows: Any, path: string): Record<string, unknown>[] {
    if (!Array.isArray(rows)) return [];
    const out: Record<string, unknown>[] = [];
    rows.forEach((row: Any, r: number) => {
      if (!isPlainObject(row)) return;
      if (row.header !== undefined) row.header = Boolean(row.header);
      const cells = Array.isArray(row.cells) ? row.cells : [];
      row.cells = cells
        .filter((cell: Any) => isPlainObject(cell) || typeof cell === "string")
        .map((cell: Any, c: number) => {
          if (typeof cell === "string") {
            return { blocks: [{ type: "paragraph", runs: [{ text: cell }] }] };
          }
          cell.blocks = this.repairBlocks(cell.blocks ?? [], `${path}/rows/${r}/cells/${c}/blocks`);
          if (Number.isInteger(cell.colSpan) && cell.colSpan > Schema.MAX_COL_SPAN) {
            this.notes.push({
              path: `${path}/rows/${r}/cells/${c}/colSpan`,
              message: `Clamped colSpan ${cell.colSpan} to ${Schema.MAX_COL_SPAN}.`,
            });
            cell.colSpan = Schema.MAX_COL_SPAN;
          }
          return cell;
        });
      out.push(row);
    });
    return out;
  }
}
