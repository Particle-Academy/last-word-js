import { isPlainObject } from "../util";
import { Schema } from "./schema";
import type { ValidationError } from "./types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function err(path: string, message: string): ValidationError {
  return { path, message };
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
        break;

      case "paragraph":
        errors.push(...this.validateRuns(block.runs, `${path}/runs`));
        if (block.align !== undefined && !(Schema.ALIGNMENTS as readonly string[]).includes(block.align)) {
          errors.push(err(`${path}/align`, "Paragraph align must be left, center, right, or justify."));
        }
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
    for (const flag of ["bold", "italic", "underline", "strike", "code"] as const) {
      if (run[flag] !== undefined && typeof run[flag] !== "boolean") {
        errors.push(err(`${path}/${flag}`, `Run \`${flag}\` must be a boolean.`));
      }
    }
    if (run.link !== undefined && typeof run.link !== "string") {
      errors.push(err(`${path}/link`, "Run `link` must be a URL string."));
    }
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
