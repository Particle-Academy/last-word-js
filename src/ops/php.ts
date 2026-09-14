/**
 * The PHP array and cast semantics the op reducer and diff depend on, in one
 * place so `doc-reducer.ts` and `doc-diff.ts` read like the PHP they mirror.
 *
 * PHP has one array type for lists and maps, and a JSON object and a JSON array
 * both decode to it. `isArr` is PHP's `is_array` over parsed JSON: any object or
 * array. A key whose value is `undefined` is treated as absent, as JSON would.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type Any = any;
export type Obj = Record<string, Any>;

/** PHP `is_array`: a JSON object or a JSON array. */
export function isArr(value: unknown): value is Obj {
  return typeof value === "object" && value !== null;
}

/** A canonical list index: "0", "1", … with no sign and no leading zero. */
function indexOf(array: readonly unknown[], key: string): number | null {
  if (!/^(0|[1-9][0-9]*)$/.test(key)) return null;
  const index = Number(key);
  return index < array.length ? index : null;
}

/**
 * PHP `array_key_exists` (a null value counts; `undefined` is absent). On a JS
 * array only an index is a key: `length` is not one.
 */
export function has(value: unknown, key: string): boolean {
  if (Array.isArray(value)) {
    const index = indexOf(value, key);
    return index !== null && value[index] !== undefined;
  }
  return isArr(value) && Object.prototype.hasOwnProperty.call(value, key) && value[key] !== undefined;
}

/** PHP `$value[$key] ?? null`. */
export function get(value: unknown, key: string): Any {
  return has(value, key) ? ((value as Obj)[key] ?? null) : null;
}

/** `foreach ($value as $key => $item)`. */
export function entriesOf(value: unknown): [string, Any][] {
  if (Array.isArray(value)) return value.map((item, i) => [String(i), item]);
  if (isArr(value)) return Object.keys(value).filter((k) => value[k] !== undefined).map((k) => [k, value[k]]);
  return [];
}

/** PHP `array_values`. */
export function valuesOf(value: unknown): Any[] {
  return entriesOf(value).map(([, item]) => item);
}

/**
 * PHP `array_is_list` over parsed JSON: a JS array, or an object keyed exactly
 * "0".."n-1" — which includes `{}`, since PHP decodes an empty object as `[]`.
 */
export function isList(value: unknown): boolean {
  if (Array.isArray(value)) return true;
  return isArr(value) && entriesOf(value).every(([key], i) => key === String(i));
}

/** PHP `ctype_digit` on a string: non-empty and ASCII digits only. */
export function isDigits(value: string): boolean {
  return /^[0-9]+$/.test(value);
}

/** PHP `(int)` cast. */
export function phpInt(value: unknown): number {
  if (value === null || value === undefined) return 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "number") return Number.isFinite(value) ? Math.trunc(value) || 0 : 0;
  if (typeof value === "string") {
    const m = /^[ \t\n\r\v\f]*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/.exec(value);
    if (!m) return 0;
    const n = Number(m[0].trim());
    return Number.isFinite(n) ? Math.trunc(n) || 0 : 0;
  }
  if (isArr(value)) return entriesOf(value).length === 0 ? 0 : 1;
  return 0;
}

/**
 * PHP `strcmp`, which `sort(SORT_STRING)` and `ksort(SORT_STRING)` use: byte
 * order of the UTF-8, which is code point order. JS `<` compares UTF-16 code
 * units, which disagrees above U+FFFF.
 */
export function compareStrings(a: string, b: string): number {
  const x = [...a];
  const y = [...b];
  for (let i = 0; i < Math.min(x.length, y.length); i++) {
    const d = x[i]!.codePointAt(0)! - y[i]!.codePointAt(0)!;
    if (d !== 0) return d < 0 ? -1 : 1;
  }
  return x.length === y.length ? 0 : x.length < y.length ? -1 : 1;
}

/** `$node[$key] = $value` on a copy. A JS array given a key that is not an index becomes a map, as a PHP list does. */
export function withKey(node: Any, key: string, value: unknown): Any {
  if (Array.isArray(node)) {
    const index = indexOf(node, key);
    if (index !== null) {
      const out = node.slice();
      out[index] = value;
      return out;
    }
    if (key === String(node.length)) {
      return [...node, value];
    }
  }
  const out: Obj = {};
  for (const [k, item] of entriesOf(node)) define(out, k, item);
  define(out, key, value);
  return out;
}

/** `unset($node[$key])` on a copy. */
export function withoutKey(node: Any, key: string): Any {
  if (!has(node, key)) return node;
  if (Array.isArray(node) && key === String(node.length - 1)) return node.slice(0, -1);
  const out: Obj = {};
  for (const [k, item] of entriesOf(node)) if (k !== key) define(out, k, item);
  return out;
}

/** Assignment that stores `__proto__` as a key, as PHP would, instead of changing the prototype. */
function define(target: Obj, key: string, value: unknown): void {
  Object.defineProperty(target, key, { value, enumerable: true, writable: true, configurable: true });
}
