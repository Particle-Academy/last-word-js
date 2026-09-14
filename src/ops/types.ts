import type { Block, Doc, ListItem, TableCell, TableRow } from "../schema/types";

/**
 * One op from `Agent.diff`, applied by `Agent.reduce`. Mirrors the PHP
 * `LastWord\Ops\DocOpSchema` variants: the same `op` names and the same fields.
 *
 * A Last Word document has no ids, so a list op names the LIST it edits with an
 * RFC 6901 JSON Pointer and the item by its 0-based index in that list:
 *
 * | op | `path` ends in | e.g. |
 * |---|---|---|
 * | `blocks.*` | `blocks` | `/blocks`, `/blocks/4/blocks` (a quote), `/blocks/2/rows/1/cells/0/blocks` |
 * | `items.*` | `items` or `children` | `/blocks/3/items`, `/blocks/3/items/0/children` |
 * | `rows.*` | `rows` | `/blocks/2/rows` |
 * | `cells.*` | `cells` | `/blocks/2/rows/1/cells` |
 */
export type DocOp =
  | DocReplaceOp
  | DocSetOp
  | BlocksInsertOp
  | BlocksRemoveOp
  | BlocksMoveOp
  | BlocksReplaceOp
  | ItemsInsertOp
  | ItemsRemoveOp
  | ItemsMoveOp
  | ItemsReplaceOp
  | RowsInsertOp
  | RowsRemoveOp
  | RowsMoveOp
  | RowsReplaceOp
  | CellsInsertOp
  | CellsRemoveOp
  | CellsMoveOp
  | CellsReplaceOp;

export type DocOpName = DocOp["op"];

/** The kinds of list an op edits. */
export type DocOpKind = "blocks" | "items" | "rows" | "cells";

/** Replace the whole document. */
export interface DocReplaceOp {
  op: "doc.replace";
  doc: Doc | Record<string, unknown>;
}

/**
 * Set a top-level property (`title`, `page`, `defaultFont`, `defaultSize`); a
 * null `value` removes it. `blocks` is refused: it changes through list ops.
 */
export interface DocSetOp {
  op: "doc.set";
  key: string;
  value: unknown;
}

/** Insert at a 0-based index, clamped to the list; a list that is not there yet is created. */
export type ListInsertOp<Kind extends DocOpKind, ValueKey extends string, Value> = {
  op: `${Kind}.insert`;
  path: string;
  index: number;
} & { [K in ValueKey]: Value };

/** Remove the item at an index; an index out of range is skipped. */
export interface ListRemoveOp<Kind extends DocOpKind> {
  op: `${Kind}.remove`;
  path: string;
  index: number;
}

/** Remove the item at `from`, then insert it at `to` (clamped); a `from` out of range is skipped. */
export interface ListMoveOp<Kind extends DocOpKind> {
  op: `${Kind}.move`;
  path: string;
  from: number;
  to: number;
}

/** Replace the item at an index; an index out of range is skipped. */
export type ListReplaceOp<Kind extends DocOpKind, ValueKey extends string, Value> = {
  op: `${Kind}.replace`;
  path: string;
  index: number;
} & { [K in ValueKey]: Value };

type AnyBlock = Block | Record<string, unknown>;
type AnyItem = ListItem | Record<string, unknown>;
type AnyRow = TableRow | Record<string, unknown>;
type AnyCell = TableCell | Record<string, unknown>;

export type BlocksInsertOp = ListInsertOp<"blocks", "block", AnyBlock>;
export type BlocksRemoveOp = ListRemoveOp<"blocks">;
export type BlocksMoveOp = ListMoveOp<"blocks">;
export type BlocksReplaceOp = ListReplaceOp<"blocks", "block", AnyBlock>;

export type ItemsInsertOp = ListInsertOp<"items", "item", AnyItem>;
export type ItemsRemoveOp = ListRemoveOp<"items">;
export type ItemsMoveOp = ListMoveOp<"items">;
export type ItemsReplaceOp = ListReplaceOp<"items", "item", AnyItem>;

export type RowsInsertOp = ListInsertOp<"rows", "row", AnyRow>;
export type RowsRemoveOp = ListRemoveOp<"rows">;
export type RowsMoveOp = ListMoveOp<"rows">;
export type RowsReplaceOp = ListReplaceOp<"rows", "row", AnyRow>;

export type CellsInsertOp = ListInsertOp<"cells", "cell", AnyCell>;
export type CellsRemoveOp = ListRemoveOp<"cells">;
export type CellsMoveOp = ListMoveOp<"cells">;
export type CellsReplaceOp = ListReplaceOp<"cells", "cell", AnyCell>;
