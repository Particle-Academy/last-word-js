import type { ValidationError } from "./schema/types";

/** Archive refusal without document names or content in the message. */
export class ArchiveRefused extends Error {
  constructor(public readonly reason: string) {
    super(`Source archive refused: ${reason}`);
    this.name = "ArchiveRefused";
    Object.setPrototypeOf(this, ArchiveRefused.prototype);
  }
}

const TEMPLATE_MESSAGES: Record<string, string> = {
  macro_enabled: "Macro-enabled templates (.dotm) are refused: re-save it as a .dotx in Word.",
  family_mismatch: "This package declares a template of a different kind than its contents.",
  already_template: "This package is already a template.",
  already_document: "This package is already a document.",
  not_convertible: "This package has no single content-type declaration that can be rewritten.",
  odf_mimetype: "An OpenDocument package must store its mimetype entry first and uncompressed.",
  bad_target: 'The conversion target must be "document" or "template".',
};

/**
 * An Office or LibreOffice template package refused, with a constant `reason`.
 * Mirrors PHP `OfficeTemplateRefused`:
 *
 * - `macro_enabled`   a .dotm (or a template carrying a VBA project). Never read
 *                     and never converted.
 * - `family_mismatch` it declares a template of another family, or its parts
 *                     are not the family it declares.
 * - `already_template` / `already_document`  conversion asked for what it is.
 * - `not_convertible` there is no single declaration that can be rewritten.
 * - `odf_mimetype`    an ODF package whose `mimetype` is not first and stored.
 * - `bad_target`      the conversion target is not `document` or `template`.
 *
 * Branch on `reason`; the message is for people.
 */
export class OfficeTemplateRefused extends Error {
  constructor(public readonly reason: string) {
    super(`Template refused (${reason}): ${TEMPLATE_MESSAGES[reason] ?? reason}`);
    this.name = "OfficeTemplateRefused";
    Object.setPrototypeOf(this, OfficeTemplateRefused.prototype);
  }
}

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
