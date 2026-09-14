import { DocReducer } from "./doc-reducer";
import type { DocOpName } from "./types";

/**
 * JSON Schema for one document op — validate ops on the wire, or register the op
 * vocabulary as an LLM tool. Mirrors PHP `LastWord\Ops\DocOpSchema`, key for key
 * (pinned byte-for-byte against the PHP package by
 * `tests/doc-ops-parity.test.ts`).
 *
 * Named like dark-slide's DeckOps (`op`, dotted `noun.verb`), with one
 * difference forced by the model: a Last Word document has no ids, so an op
 * addresses a list by JSON Pointer and an item by index in it. See `DocReducer`
 * for what each path may point at.
 */
export const DocOpSchema = {
  /** Every op name, in the order the variants are listed. */
  TYPES: [
    "doc.replace",
    "doc.set",
    "blocks.insert",
    "blocks.remove",
    "blocks.move",
    "blocks.replace",
    "items.insert",
    "items.remove",
    "items.move",
    "items.replace",
    "rows.insert",
    "rows.remove",
    "rows.move",
    "rows.replace",
    "cells.insert",
    "cells.remove",
    "cells.move",
    "cells.replace",
  ] as const satisfies readonly DocOpName[],

  /** A fresh object on every call, so a caller editing it changes nothing here. */
  jsonSchema(): Record<string, unknown> {
    const index = () => ({ type: "integer", minimum: 0 });
    const variants = [
      variant("doc.replace", { doc: { type: "object" } }, ["doc"], "Replace the whole document."),
      variant(
        "doc.set",
        { key: { type: "string", minLength: 1, not: { const: "blocks" } }, value: { description: "Any JSON value; null removes the key." } },
        ["key", "value"],
        "Set a top-level property (title, page, defaultFont, defaultSize); null removes it.",
      ),
    ];

    for (const [kind, [valueKey, ends]] of Object.entries(DocReducer.KINDS)) {
      const path = () => ({ type: "string", pattern: "^(/[^/]+)*/(" + ends.join("|") + ")$" });
      const value = () => ({ type: "object" });

      variants.push(
        variant(`${kind}.insert`, { path: path(), index: index(), [valueKey]: value() }, ["path", "index", valueKey], `Insert a ${valueKey} at a 0-based index in the list at path.`),
        variant(`${kind}.remove`, { path: path(), index: index() }, ["path", "index"], `Remove the ${valueKey} at an index in the list at path.`),
        variant(`${kind}.move`, { path: path(), from: index(), to: index() }, ["path", "from", "to"], `Move a ${valueKey} within the list at path: removed at from, inserted at to.`),
        variant(`${kind}.replace`, { path: path(), index: index(), [valueKey]: value() }, ["path", "index", valueKey], `Replace the ${valueKey} at an index in the list at path.`),
      );
    }

    return {
      $schema: "http://json-schema.org/draft-07/schema#",
      title: "Last Word op",
      description: "One op from Agent::diff, applied by Agent::reduce.",
      oneOf: variants,
    };
  },
};

/** Fails `tsc` if the `DocOp` union names an op that `TYPES` does not list. */
type Unlisted = Exclude<DocOpName, (typeof DocOpSchema.TYPES)[number]>;
const DOC_OP_TYPES_COVER_THE_UNION: [Unlisted] extends [never] ? true : Unlisted = true;

function variant(op: string, properties: Record<string, unknown>, required: string[], description: string): Record<string, unknown> {
  return {
    type: "object",
    description,
    required: ["op", ...required],
    additionalProperties: false,
    properties: { op: { const: op }, ...properties },
  };
}
