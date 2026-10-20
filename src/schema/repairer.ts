import { clone, isNumeric, isPlainObject } from "../util";
import { Schema } from "./schema";
import type { ValidationError } from "./types";
import { validateChart } from "./validator";
import { chartLabel } from "../helpers/chart";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

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

  repair(docInput: Any): Record<string, unknown> {
    this.notes = [];
    const doc: Any = isPlainObject(docInput) ? clone(docInput) : {};

    if (doc.title !== undefined && typeof doc.title !== "string") {
      doc.title = String(doc.title);
    }
    doc.blocks = this.repairBlocks(doc.blocks ?? [], "/blocks");

    return doc;
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
        block.runs = this.repairRuns(block.runs);
        break;
      }
      case "paragraph":
        block.runs = this.repairRuns(block.runs);
        if (block.align !== undefined && !(Schema.ALIGNMENTS as readonly string[]).includes(block.align)) {
          delete block.align;
        }
        break;

      case "list":
        if (block.ordered !== undefined) block.ordered = Boolean(block.ordered);
        block.items = this.repairListItems(block.items);
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

  private repairRuns(runs: Any): Record<string, unknown>[] {
    // A bare string coerces to a single run.
    if (typeof runs === "string") return [{ text: runs }];
    if (!Array.isArray(runs)) return [];
    const out: Record<string, unknown>[] = [];
    for (const run of runs) {
      if (typeof run === "string") {
        out.push({ text: run });
        continue;
      }
      if (!isPlainObject(run)) continue;
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

  private repairListItems(items: Any): Record<string, unknown>[] {
    if (!Array.isArray(items)) return [];
    const out: Record<string, unknown>[] = [];
    for (const item of items) {
      if (typeof item === "string") {
        out.push({ runs: [{ text: item }] });
        continue;
      }
      if (!isPlainObject(item)) continue;
      item.runs = this.repairRuns(item.runs);
      if (item.children !== undefined) {
        const children = this.repairListItems(item.children);
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
          return cell;
        });
      out.push(row);
    });
    return out;
  }
}
