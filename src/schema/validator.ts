import { isPlainObject } from "../util";
import { Schema } from "./schema";
import type { ValidationError } from "./types";
import { SUPPORTED_CHART_TYPES, pointValue } from "../helpers/chart";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function err(path: string, message: string): ValidationError {
  return { path, message };
}

const HEX = /^#[0-9a-fA-F]{6}$/;
const BORDER_STYLES = ["single", "double", "dashed", "dotted", "none"];
const BOX_EDGES = ["top", "right", "bottom", "left"];
const TABLE_EDGES = [...BOX_EDGES, "insideH", "insideV"];

function isNum(v: Any): boolean {
  return typeof v === "number" && Number.isFinite(v);
}

/** A finite number, or absent. */
function checkNumber(value: Any, path: string, label: string, positive = false): ValidationError[] {
  if (value === undefined) return [];
  if (!isNum(value) || (positive && value <= 0)) {
    return [err(path, `${label} must be a ${positive ? "positive " : ""}number of points.`)];
  }
  return [];
}

function checkBoolean(value: Any, path: string, label: string): ValidationError[] {
  return value === undefined || typeof value === "boolean" ? [] : [err(path, `${label} must be a boolean.`)];
}

function checkHex(value: Any, path: string, label: string): ValidationError[] {
  return value === undefined || (typeof value === "string" && HEX.test(value))
    ? []
    : [err(path, `${label} must be a #RRGGBB hex string.`)];
}

function checkEnum(value: Any, path: string, label: string, allowed: string[]): ValidationError[] {
  return value === undefined || (typeof value === "string" && allowed.includes(value))
    ? []
    : [err(path, `${label} must be one of ${allowed.join(", ")}.`)];
}

/** `{ top?, right?, bottom?, left? }` of points. */
function checkBoxSides(value: Any, path: string, label: string): ValidationError[] {
  if (value === undefined) return [];
  if (!isPlainObject(value)) return [err(path, `${label} must be an object of sides in points.`)];
  const errors: ValidationError[] = [];
  for (const side of BOX_EDGES) {
    errors.push(...checkNumber(value[side], `${path}/${side}`, `\`${side}\``));
  }
  return errors;
}

function checkBorders(value: Any, path: string, label: string, edges: string[]): ValidationError[] {
  if (value === undefined) return [];
  if (!isPlainObject(value)) return [err(path, `${label} must be an object keyed by edge.`)];
  const errors: ValidationError[] = [];
  for (const edge of edges) {
    const border = value[edge];
    if (border === undefined) continue;
    const edgePath = `${path}/${edge}`;
    if (!isPlainObject(border)) {
      errors.push(err(edgePath, "A border must be an object (`{ style?, width?, color? }`)."));
      continue;
    }
    errors.push(...checkEnum(border.style, `${edgePath}/style`, "Border `style`", BORDER_STYLES));
    errors.push(...checkNumber(border.width, `${edgePath}/width`, "Border `width`", true));
    errors.push(...checkHex(border.color, `${edgePath}/color`, "Border `color`"));
  }
  return errors;
}

/**
 * A break asking for a different page. A .docx has one page setup until the
 * writer implements `page`, so it is refused rather than ignored: a document that
 * silently comes out all portrait looks fine. Repair leaves these keys in place,
 * so `validateAndRepair` stays `ok: false`. The messages are PHP's, byte for byte
 * (unlike the rest of this file, which has its own sentence case), so all three
 * engines say the same thing about this.
 */
function validatePageBreak(block: Any, path: string): ValidationError[] {
  const errors: ValidationError[] = [];
  if (block.page !== undefined) {
    errors.push(
      err(
        `${path}/page`,
        "pageBreak.page is not implemented in last-word yet; remove page to keep one page setup, or write this document as a PDF with fine-print",
      ),
    );
  }
  for (const key of Schema.PAGE_BREAK_NEAR_MISS_KEYS) {
    if (block[key] !== undefined) {
      errors.push(err(`${path}/${key}`, `a pageBreak has no ${key} key; use pageBreak.page`));
    }
  }
  return errors;
}

/**
 * The properties every paragraph-shaped block accepts.
 *
 * Shared by `paragraph`, `heading` and list items, because a heading is a
 * paragraph — before this, a section label that needed spacing had to be a bold
 * paragraph impersonating one.
 */
function checkParagraphProps(block: Any, path: string): ValidationError[] {
  return [
    ...checkNumber(block.spaceBefore, `${path}/spaceBefore`, "`spaceBefore`"),
    ...checkNumber(block.spaceAfter, `${path}/spaceAfter`, "`spaceAfter`"),
    ...checkNumber(block.lineHeight, `${path}/lineHeight`, "`lineHeight` (a multiple of single spacing)", true),
    ...checkNumber(block.indentLeft, `${path}/indentLeft`, "`indentLeft`"),
    ...checkNumber(block.indentRight, `${path}/indentRight`, "`indentRight`"),
    ...checkBoolean(block.keepNext, `${path}/keepNext`, "`keepNext`"),
    ...checkHex(block.shading, `${path}/shading`, "`shading`"),
    ...checkBorders(block.borders, `${path}/borders`, "`borders`", BOX_EDGES),
  ];
}

/**
 * A chart block. Everything `chartSpec()` needs to draw is checked here, so a
 * valid chart is always a drawable one. Mirrors PHP `Validator::validateChart()`
 * (null counts as absent, as `isset` does there).
 */
export function validateChart(block: Any, path: string): ValidationError[] {
  const errors: ValidationError[] = [];
  const present = (v: unknown): boolean => v !== undefined && v !== null;

  for (const dim of ["widthPx", "heightPx"] as const) {
    if (present(block[dim]) && (!isNum(block[dim]) || block[dim] <= 0)) {
      errors.push(err(`${path}/${dim}`, `Chart \`${dim}\` must be a positive number of px.`));
    }
  }
  if (present(block.width) && (!isNum(block.width) || block.width <= 0 || block.width > 100)) {
    errors.push(err(`${path}/width`, "Chart `width` must be a percentage of the text column, above 0 and at most 100."));
  }
  if (present(block.numberFormat) && (typeof block.numberFormat !== "string" || block.numberFormat === "")) {
    errors.push(err(`${path}/numberFormat`, 'Chart `numberFormat` must be an Excel number format string, e.g. "#,##0" or "0.0%".'));
  }
  if (present(block.alt) && typeof block.alt !== "string") {
    errors.push(err(`${path}/alt`, "Chart `alt` must be a string."));
  }

  const option = block.option;
  if (!isPlainObject(option)) {
    errors.push(
      err(
        `${path}/option`,
        'Chart must have an `option` object: an Apache ECharts option, e.g. {"xAxis": {"type": "category", "data": ["Q1", "Q2"]}, "yAxis": {"type": "value"}, "series": [{"type": "bar", "name": "Revenue", "data": [120, 135]}]}.',
      ),
    );
    return errors;
  }
  const optPath = `${path}/option`;

  let series: Any = (option as Any).series;
  if (isPlainObject(series)) series = [series];
  if (!Array.isArray(series) || series.length === 0) {
    errors.push(err(`${optPath}/series`, "Chart option must have `series`: a list of {type, name, data} objects."));
    return errors;
  }

  const supported = SUPPORTED_CHART_TYPES.join(", ");
  let kind: string | null = null;
  series.forEach((one: Any, i: number) => {
    const serPath = `${optPath}/series/${i}`;
    if (!isPlainObject(one)) {
      errors.push(err(serPath, "A series must be an object ({type, name, data})."));
      return;
    }
    const type = one.type ?? "bar";
    if (typeof type !== "string" || !(SUPPORTED_CHART_TYPES as readonly string[]).includes(type)) {
      errors.push(err(`${serPath}/type`, `Series \`type\` must be one of ${supported} (a document chart cannot draw ${typeof type === "string" ? `"${type}"` : "that"}).`));
      return;
    }
    if (kind === null) kind = type;
    else if (type !== kind) {
      errors.push(err(`${serPath}/type`, `Every series in a chart must be the same type: this one is "${type}", the first is "${kind}".`));
    }
    if (present(one.name) && typeof one.name !== "string" && !isNum(one.name)) {
      errors.push(err(`${serPath}/name`, "Series `name` must be a string."));
    }
    errors.push(...checkHexPresent(isPlainObject(one.itemStyle) ? one.itemStyle.color : undefined, `${serPath}/itemStyle/color`, "Series colour"));

    const data = one.data;
    if (!Array.isArray(data)) {
      const hint = one.values !== undefined ? ' (ECharts calls it "data", not "values")' : "";
      errors.push(err(`${serPath}/data`, `Series must have a \`data\` list of numbers${hint}.`));
      return;
    }
    data.forEach((point: Any, p: number) => {
      if (pointValue(point) === null) {
        errors.push(
          err(`${serPath}/data/${p}`, type === "pie" ? 'A pie point must be {"name": string, "value": number}.' : 'A point must be a number, or {"value": number}.'),
        );
        return;
      }
      if (isPlainObject(point)) {
        errors.push(...checkHexPresent(isPlainObject(point.itemStyle) ? point.itemStyle.color : undefined, `${serPath}/data/${p}/itemStyle/color`, "Point colour"));
      }
    });
  });
  if (kind === "pie" && series.length > 1) {
    errors.push(err(`${optPath}/series`, "A pie chart takes exactly one series."));
  }

  for (const key of ["xAxis", "yAxis"] as const) {
    let axis: Any = (option as Any)[key];
    if (!present(axis)) continue;
    if (Array.isArray(axis) && axis.length > 0) axis = axis[0];
    if (!isPlainObject(axis)) {
      errors.push(err(`${optPath}/${key}`, `\`${key}\` must be an object ({type, data, name}).`));
      continue;
    }
    if (present(axis.type)) errors.push(...checkEnum(axis.type, `${optPath}/${key}/type`, `\`${key}.type\``, ["category", "value"]));
    if (present(axis.data) && !Array.isArray(axis.data)) {
      errors.push(err(`${optPath}/${key}/data`, `\`${key}.data\` must be a list of category labels.`));
    }
    if (present(axis.name) && typeof axis.name !== "string") {
      errors.push(err(`${optPath}/${key}/name`, `\`${key}.name\` (the axis title) must be a string.`));
    }
  }

  const color = (option as Any).color;
  if (present(color)) {
    if (!Array.isArray(color)) {
      errors.push(err(`${optPath}/color`, "`option.color` must be a list of #RRGGBB colours."));
    } else {
      color.forEach((c: Any, i: number) => errors.push(...checkHexPresent(c, `${optPath}/color/${i}`, "Palette colour")));
    }
  }

  return errors;
}

/** A #RRGGBB colour, or absent (null counts as absent). */
function checkHexPresent(value: Any, path: string, label: string): ValidationError[] {
  return value === null ? [] : checkHex(value, path, label);
}

/** How deep the footnote refusal walk goes, as PHP's `Portable::MAX_DEPTH`. */
const MAX_DEPTH = 64;

/** True when a run carries a `footnote` key (any value), as PHP's `array_key_exists`. */
export function hasFootnote(run: Any): boolean {
  return isPlainObject(run) && Object.prototype.hasOwnProperty.call(run, "footnote") && (run as Any).footnote !== undefined;
}

/**
 * Every run, with its path, in a list of blocks: paragraphs, headings, list
 * items at any depth, table cells and quotes. Not into a footnote's own blocks.
 */
export function eachRun(blocks: Any, path: string, visit: (run: Any, runPath: string) => void, depth = 0): void {
  if (!Array.isArray(blocks) || depth > MAX_DEPTH) return;
  blocks.forEach((block: Any, i: number) => {
    if (!isPlainObject(block)) return;
    const at = `${path}/${i}`;
    const b = block as Any;
    const runs = (list: Any, listPath: string): void => {
      if (Array.isArray(list)) list.forEach((run: Any, r: number) => visit(run, `${listPath}/${r}`));
    };
    const items = (list: Any, listPath: string, d: number): void => {
      if (!Array.isArray(list) || d > MAX_DEPTH) return;
      list.forEach((item: Any, n: number) => {
        if (!isPlainObject(item)) return;
        runs((item as Any).runs, `${listPath}/${n}/runs`);
        items((item as Any).children, `${listPath}/${n}/children`, d + 1);
      });
    };
    if (b.type === "paragraph" || b.type === "heading") runs(b.runs, `${at}/runs`);
    if (b.type === "list") items(b.items, `${at}/items`, depth + 1);
    if (b.type === "quote") eachRun(b.blocks, `${at}/blocks`, visit, depth + 1);
    if (b.type === "table" && Array.isArray(b.rows)) {
      b.rows.forEach((row: Any, r: number) => {
        if (!isPlainObject(row) || !Array.isArray((row as Any).cells)) return;
        (row as Any).cells.forEach((cell: Any, c: number) => {
          if (isPlainObject(cell)) eachRun((cell as Any).blocks, `${at}/rows/${r}/cells/${c}/blocks`, visit, depth + 1);
        });
      });
    }
  });
}

/** Liberal schema validator. Mirrors PHP `Schema\Validator`. */
export class Validator {
  /**
   * Where the blocks being validated live: the body, or a footnote. A footnote
   * run is refused inside a footnote -- Word has none inside another -- and in
   * a header or footer (see the band check in `validate()`).
   */
  private context: "body" | "footnote" = "body";

  validate(doc: Any): ValidationError[] {
    this.context = "body";
    const errors: ValidationError[] = [];
    if (!isPlainObject(doc)) {
      return [err("/", "Doc must be a JSON object with a `blocks` array.")];
    }

    if (doc.sections !== undefined) {
      errors.push(err("/sections", "there is no sections key; use pageBreak.page"));
    }

    if (doc.title !== undefined && typeof doc.title !== "string") {
      errors.push(err("/title", "Doc title must be a string."));
    }

    // `meta` (PHP 0.16.0): an object whose `author` is text. The title is the
    // top-level `title`, not `meta.title` -- one slot for dc:title, so the two
    // cannot disagree; repair moves a `meta.title` there. An empty JSON array
    // is PHP's empty object too.
    if (doc.meta !== undefined) {
      const meta = doc.meta;
      if (!isPlainObject(meta) && !(Array.isArray(meta) && meta.length === 0)) {
        errors.push(err("/meta", "Doc `meta` must be an object ({author?})."));
      } else if (isPlainObject(meta)) {
        if ("author" in meta && typeof (meta as Any).author !== "string") {
          errors.push(err("/meta/author", "Doc `meta.author` must be text (docProps/core.xml dc:creator)."));
        }
        if ("title" in meta) {
          errors.push(err("/meta/title", 'The document title is the top-level "title", not meta.title.'));
        }
      }
    }

    if (doc.defaultFont !== undefined && typeof doc.defaultFont !== "string") {
      errors.push(err("/defaultFont", "Doc `defaultFont` must be a font family name."));
    }
    errors.push(...checkNumber(doc.defaultSize, "/defaultSize", "Doc `defaultSize`", true));

    if (doc.page !== undefined) {
      if (!isPlainObject(doc.page)) {
        errors.push(err("/page", "Doc `page` must be an object (`{ size?, orientation?, margins? }`)."));
      } else {
        errors.push(...checkEnum(doc.page.size, "/page/size", "Page `size`", ["letter", "legal", "a4"]));
        errors.push(
          ...checkEnum(doc.page.orientation, "/page/orientation", "Page `orientation`", ["portrait", "landscape"]),
        );
        errors.push(...checkBoxSides(doc.page.margins, "/page/margins", "Page `margins`"));
        const margins = doc.page.margins;
        for (const distance of ["header", "footer"] as const) {
          const value = isPlainObject(margins) ? (margins as Any)[distance] : undefined;
          if (value !== undefined && value !== null && (!isNum(value) || value < 0)) {
            errors.push(
              err(
                `/page/margins/${distance}`,
                `Page margins \`${distance}\` is the distance from the page edge to the ${distance}: a number of points, 0 or more (36 by default).`,
              ),
            );
          }
        }
        for (const kind of Schema.PAGE_BANDS) {
          const band = doc.page[kind];
          if (band === undefined || band === null) continue;
          if (!isPlainObject(band)) {
            errors.push(err(`/page/${kind}`, `Page \`${kind}\` must be an object (\`{ left?, center?, right?, size?, color? }\`).`));
            continue;
          }
          // A header or footer cannot hold a footnote. This engine writes no
          // `{blocks}` band (PHP 0.14.0, last-word#6, is not ported), but an
          // agent writing one for either engine is told the same thing here.
          eachRun((band as Any).blocks, `/page/${kind}/blocks`, (run, runPath) => {
            if (hasFootnote(run)) errors.push(err(`${runPath}/footnote`, "A header or footer cannot hold a footnote."));
          });
          for (const slot of ["left", "center", "right"]) {
            if (band[slot] !== undefined && typeof band[slot] !== "string") {
              errors.push(err(`/page/${kind}/${slot}`, `Page \`${kind}.${slot}\` must be text.`));
            }
          }
          if (band.size !== undefined && !(typeof band.size === "number" && band.size > 0)) {
            errors.push(err(`/page/${kind}/size`, `Page \`${kind}.size\` must be a positive number of points.`));
          }
          if (band.color !== undefined && !(typeof band.color === "string" && /^#[0-9A-Fa-f]{6}$/.test(band.color))) {
            errors.push(err(`/page/${kind}/color`, `Page \`${kind}.color\` must be #RRGGBB.`));
          }
        }
      }
    }

    if (!("blocks" in doc)) {
      errors.push(err("/blocks", "Doc must have a `blocks` array."));
    } else if (!Array.isArray(doc.blocks)) {
      errors.push(err("/blocks", "Doc blocks must be a JSON array."));
    } else {
      doc.blocks.forEach((block: Any, i: number) => {
        errors.push(...this.validateBlock(block, `/blocks/${i}`));
      });
    }

    return errors;
  }

  validateBlock(block: Any, path: string): ValidationError[] {
    const errors: ValidationError[] = [];
    if (!isPlainObject(block)) {
      return [err(path, "Each block must be a JSON object with a `type` field.")];
    }
    if (typeof block.type !== "string") {
      return [err(`${path}/type`, "Block must have a string `type` field.")];
    }
    if (!(Schema.BLOCK_TYPES as readonly string[]).includes(block.type)) {
      return [err(`${path}/type`, `Unknown block type \`${block.type}\`.`)];
    }

    switch (block.type) {
      case "heading":
        if (!Number.isInteger(block.level) || block.level < 1 || block.level > Schema.MAX_HEADING_LEVEL) {
          errors.push(err(`${path}/level`, `Heading level must be an integer between 1 and ${Schema.MAX_HEADING_LEVEL}.`));
        }
        errors.push(...this.validateRuns(block.runs, `${path}/runs`));
        errors.push(...this.validateAlign(block, path));
        errors.push(...checkParagraphProps(block, path));
        break;

      case "paragraph":
        errors.push(...this.validateRuns(block.runs, `${path}/runs`));
        errors.push(...this.validateAlign(block, path));
        errors.push(...checkParagraphProps(block, path));
        break;

      case "list":
        if (block.ordered !== undefined && typeof block.ordered !== "boolean") {
          errors.push(err(`${path}/ordered`, "List `ordered` must be a boolean."));
        }
        if (!Array.isArray(block.items)) {
          errors.push(err(`${path}/items`, "List must have an `items` array."));
        } else {
          block.items.forEach((item: Any, i: number) => {
            errors.push(...this.validateListItem(item, `${path}/items/${i}`));
          });
        }
        break;

      case "table":
        errors.push(...checkEnum(block.align, `${path}/align`, "Table `align`", ["left", "center", "right"]));
        errors.push(...checkBorders(block.borders, `${path}/borders`, "Table `borders`", TABLE_EDGES));
        errors.push(...checkBoxSides(block.cellPadding, `${path}/cellPadding`, "Table `cellPadding`"));
        if (block.width !== undefined && (!isNum(block.width) || block.width <= 0 || block.width > 100)) {
          errors.push(err(`${path}/width`, "Table `width` must be a percentage of the text column, above 0 and at most 100."));
        }
        if (block.widths !== undefined) {
          if (!Array.isArray(block.widths) || block.widths.length === 0) {
            errors.push(err(`${path}/widths`, "Table `widths` must be a non-empty array of relative column weights."));
          } else {
            block.widths.forEach((w: Any, i: number) => {
              if (!isNum(w) || w < 0) {
                errors.push(err(`${path}/widths/${i}`, "A column weight must be a non-negative number."));
              }
            });
          }
        }
        if (!Array.isArray(block.rows)) {
          errors.push(err(`${path}/rows`, "Table must have a `rows` array."));
        } else {
          block.rows.forEach((row: Any, r: number) => {
            const rowPath = `${path}/rows/${r}`;
            if (!isPlainObject(row)) {
              errors.push(err(rowPath, "Each table row must be a JSON object with a `cells` array."));
              return;
            }
            if (row.header !== undefined && typeof row.header !== "boolean") {
              errors.push(err(`${rowPath}/header`, "Row `header` must be a boolean."));
            }
            if (!Array.isArray(row.cells)) {
              errors.push(err(`${rowPath}/cells`, "Table row must have a `cells` array."));
              return;
            }
            row.cells.forEach((cell: Any, c: number) => {
              const cellPath = `${rowPath}/cells/${c}`;
              if (!isPlainObject(cell)) {
                errors.push(err(cellPath, "Each table cell must be a JSON object with a `blocks` array."));
                return;
              }
              errors.push(...checkHex(cell.shading, `${cellPath}/shading`, "Cell `shading`"));
              errors.push(...checkBorders(cell.borders, `${cellPath}/borders`, "Cell `borders`", BOX_EDGES));
              errors.push(...checkBoxSides(cell.padding, `${cellPath}/padding`, "Cell `padding`"));
              errors.push(...checkEnum(cell.valign, `${cellPath}/valign`, "Cell `valign`", ["top", "center", "bottom"]));
              for (const span of ["colSpan", "rowSpan"] as const) {
                if (cell[span] !== undefined && (!Number.isInteger(cell[span]) || cell[span] < 1)) {
                  errors.push(err(`${cellPath}/${span}`, `Cell \`${span}\` must be an integer of 1 or more.`));
                }
              }
              if (!Array.isArray(cell.blocks)) {
                errors.push(err(`${cellPath}/blocks`, "Table cell must have a `blocks` array."));
                return;
              }
              cell.blocks.forEach((b: Any, i: number) => {
                errors.push(...this.validateBlock(b, `${cellPath}/blocks/${i}`));
              });
            });
          });
        }
        break;

      case "code":
        if (typeof block.text !== "string") {
          errors.push(err(`${path}/text`, "Code block must have a string `text` field."));
        }
        if (block.language !== undefined && typeof block.language !== "string") {
          errors.push(err(`${path}/language`, "Code block `language` must be a string."));
        }
        break;

      case "quote":
        if (!Array.isArray(block.blocks)) {
          errors.push(err(`${path}/blocks`, "Quote must have a `blocks` array."));
        } else {
          block.blocks.forEach((b: Any, i: number) => {
            errors.push(...this.validateBlock(b, `${path}/blocks/${i}`));
          });
        }
        break;

      case "image":
        if (typeof block.src !== "string" || block.src === "") {
          errors.push(err(`${path}/src`, "Image must have a `src` data URL string."));
        } else if (!block.src.startsWith("data:image/")) {
          errors.push(err(`${path}/src`, "Image `src` must be a data:image/png or data:image/jpeg data URL."));
        }
        for (const key of ["widthPx", "heightPx"] as const) {
          if (block[key] !== undefined && (typeof block[key] !== "number" || !(block[key] > 0))) {
            errors.push(err(`${path}/${key}`, `Image \`${key}\` must be a positive number.`));
          }
        }
        if (block.alt !== undefined && typeof block.alt !== "string") {
          errors.push(err(`${path}/alt`, "Image `alt` must be a string."));
        }
        break;

      case "chart":
        errors.push(...validateChart(block, path));
        break;

      case "pageBreak":
        errors.push(...validatePageBreak(block, path));
        break;

      case "hr":
        break;
    }

    return errors;
  }

  private validateAlign(block: Any, path: string): ValidationError[] {
    return block.align === undefined || (Schema.ALIGNMENTS as readonly string[]).includes(block.align)
      ? []
      : [err(`${path}/align`, "Align must be left, center, right, or justify.")];
  }

  private validateRuns(runs: Any, path: string): ValidationError[] {
    const errors: ValidationError[] = [];
    if (!Array.isArray(runs)) {
      return [err(path, "Must be an array of runs (`{ text, bold?, italic?, … }`).")];
    }
    runs.forEach((run: Any, i: number) => {
      errors.push(...this.validateRun(run, `${path}/${i}`));
    });
    return errors;
  }

  private validateRun(run: Any, path: string): ValidationError[] {
    const errors: ValidationError[] = [];
    if (!isPlainObject(run)) {
      return [err(path, "Each run must be a JSON object with a `text` field.")];
    }
    const isNote = hasFootnote(run);
    if (typeof run.text !== "string") {
      errors.push(
        err(`${path}/text`, isNote ? 'A footnote run requires `"text": ""` (Word numbers the mark).' : "Run must have a string `text` field."),
      );
    }
    if (isNote) errors.push(...this.validateFootnote(run, path));
    for (const flag of ["bold", "italic", "underline", "strike", "code", "smallCaps"] as const) {
      if (run[flag] !== undefined && typeof run[flag] !== "boolean") {
        errors.push(err(`${path}/${flag}`, `Run \`${flag}\` must be a boolean.`));
      }
    }
    if (run.link !== undefined && typeof run.link !== "string") {
      errors.push(err(`${path}/link`, "Run `link` must be a URL string."));
    }
    if (run.font !== undefined && typeof run.font !== "string") {
      errors.push(err(`${path}/font`, "Run `font` must be a font family name."));
    }
    errors.push(...checkNumber(run.size, `${path}/size`, "Run `size`", true));
    errors.push(...checkNumber(run.letterSpacing, `${path}/letterSpacing`, "Run `letterSpacing`"));
    for (const key of ["color", "highlight"] as const) {
      if (run[key] !== undefined && (typeof run[key] !== "string" || !/^#[0-9a-fA-F]{6}$/.test(run[key]))) {
        errors.push(err(`${path}/${key}`, `Run \`${key}\` must be a #RRGGBB hex string.`));
      }
    }
    return errors;
  }

  /**
   * A footnote run: `{"text": "", "footnote": {"blocks": [...]}}`, its blocks
   * the body's rules less `Schema.FOOTNOTE_REFUSED_TYPES` at any depth. Mirrors
   * PHP `Validator::validateFootnote()`.
   */
  private validateFootnote(run: Any, runPath: string): ValidationError[] {
    const at = `${runPath}/footnote`;
    if (this.context === "footnote") {
      return [err(at, "A footnote cannot hold a footnote.")];
    }
    const errors: ValidationError[] = [];
    if (typeof run.text === "string" && run.text !== "") {
      errors.push(err(`${runPath}/text`, 'A footnote run\'s text must be "": Word numbers the mark. Put the text it follows in the run before it.'));
    }
    const note = run.footnote;
    // An empty JSON array is PHP's empty object too: both are "no blocks", below.
    if (!isPlainObject(note) && !(Array.isArray(note) && note.length === 0)) {
      errors.push(err(at, 'Footnote must be an object: {"blocks": [{"type": "paragraph", "runs": [{"text": "..."}]}]}.'));
      return errors;
    }
    const blocks = (note as Any).blocks;
    if (!Array.isArray(blocks) || blocks.length === 0) {
      errors.push(
        err(`${at}/blocks`, 'Footnote requires `blocks`: a non-empty list of blocks, e.g. [{"type": "paragraph", "runs": [{"text": "Source: ..."}]}].'),
      );
      return errors;
    }

    const outer = this.context;
    this.context = "footnote";
    try {
      blocks.forEach((block: Any, i: number) => {
        if (isPlainObject(block) && (Schema.FOOTNOTE_REFUSED_TYPES as readonly string[]).includes((block as Any).type)) {
          return; // refused below
        }
        errors.push(...this.validateBlock(block, `${at}/blocks/${i}`));
      });
    } finally {
      this.context = outer;
    }
    this.refuseInFootnote(blocks, `${at}/blocks`, errors);
    return errors;
  }

  private refuseInFootnote(blocks: Any[], path: string, errors: ValidationError[], depth = 0): void {
    if (depth > MAX_DEPTH) return;
    blocks.forEach((block: Any, i: number) => {
      if (!isPlainObject(block)) return;
      const at = `${path}/${i}`;
      const type = (block as Any).type;
      if ((Schema.FOOTNOTE_REFUSED_TYPES as readonly string[]).includes(type)) {
        errors.push(err(`${at}/type`, `A footnote cannot hold a ${type} block.`));
        return;
      }
      if (type === "quote" && Array.isArray((block as Any).blocks)) {
        this.refuseInFootnote((block as Any).blocks, `${at}/blocks`, errors, depth + 1);
      }
      if (type === "table" && Array.isArray((block as Any).rows)) {
        (block as Any).rows.forEach((row: Any, r: number) => {
          const cells = isPlainObject(row) && Array.isArray((row as Any).cells) ? (row as Any).cells : [];
          cells.forEach((cell: Any, c: number) => {
            if (isPlainObject(cell) && Array.isArray((cell as Any).blocks)) {
              this.refuseInFootnote((cell as Any).blocks, `${at}/rows/${r}/cells/${c}/blocks`, errors, depth + 1);
            }
          });
        });
      }
    });
  }

  private validateListItem(item: Any, path: string): ValidationError[] {
    const errors: ValidationError[] = [];
    if (!isPlainObject(item)) {
      return [err(path, "Each list item must be a JSON object with a `runs` array.")];
    }
    errors.push(...this.validateRuns(item.runs, `${path}/runs`));
    errors.push(...checkParagraphProps(item, path));
    if (item.children !== undefined) {
      if (!Array.isArray(item.children)) {
        errors.push(err(`${path}/children`, "List item `children` must be an array of list items."));
      } else {
        item.children.forEach((child: Any, i: number) => {
          errors.push(...this.validateListItem(child, `${path}/children/${i}`));
        });
      }
    }
    return errors;
  }
}
