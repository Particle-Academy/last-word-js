import { isPlainObject } from "../util";
import { Schema } from "./schema";
import type { ValidationError } from "./types";

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

/** Liberal schema validator. Mirrors PHP `Schema\Validator`. */
export class Validator {
  validate(doc: Any): ValidationError[] {
    const errors: ValidationError[] = [];
    if (!isPlainObject(doc)) {
      return [err("/", "Doc must be a JSON object with a `blocks` array.")];
    }

    if (doc.title !== undefined && typeof doc.title !== "string") {
      errors.push(err("/title", "Doc title must be a string."));
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

      case "pageBreak":
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
    if (typeof run.text !== "string") {
      errors.push(err(`${path}/text`, "Run must have a string `text` field."));
    }
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
