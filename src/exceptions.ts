import type { ValidationError } from "./schema/types";

/** Thrown by Agent.write/toBytes when the schema is invalid. Mirrors PHP `SchemaException`. */
export class SchemaException extends Error {
  readonly errors: ValidationError[];

  constructor(message: string, errors: ValidationError[]) {
    super(message);
    this.name = "SchemaException";
    this.errors = errors;
    Object.setPrototypeOf(this, SchemaException.prototype);
  }
}

/**
 * The bytes are a document we recognise and cannot read. Mirrors PHP
 * `UnsupportedFormatException`.
 *
 * Deliberately distinct from the plain `Error` a damaged file raises, because
 * the two need different things said to a person: "this file is damaged" and
 * "this is a format we do not read, save it as .docx" lead to different actions.
 *
 * `format` carries the detected format so a host can branch without parsing the
 * message: `doc`, `xls`, `ppt`, `msg`, `cfb`, `xlsx`, `pptx`, `ods`, `odp`,
 * `unknown`.
 */
export class UnsupportedFormatException extends Error {
  readonly format: string;

  constructor(format: string, message: string) {
    super(message);
    this.name = "UnsupportedFormatException";
    this.format = format;
    Object.setPrototypeOf(this, UnsupportedFormatException.prototype);
  }
}

/**
 * A `template` passed to `Agent.toBytes()` / `Agent.write()` cannot be used —
 * not a zip, or no `word/styles.xml`. Mirrors PHP `TemplateException`.
 *
 * It REFUSES rather than falling back to the built-in look, which is the point
 * of last-word#3: a document that silently comes out in the default style is
 * exactly the failure the template option exists to end, and a host that asked
 * for a house template would rather hear why than ship the wrong file. It also
 * lets a host validate a customer-supplied template at upload time by trying to
 * open it.
 */
export class TemplateException extends Error {
  /** The underlying failure, when one caused this (e.g. a zip that would not inflate). */
  readonly reason?: unknown;

  constructor(message: string, reason?: unknown) {
    super(message);
    this.name = "TemplateException";
    this.reason = reason;
    Object.setPrototypeOf(this, TemplateException.prototype);
  }
}
