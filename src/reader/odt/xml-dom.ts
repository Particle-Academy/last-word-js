/**
 * A small XML parser that keeps what an ODF reader needs and `parseXml` drops:
 * the ORDER of text and elements inside a paragraph (`a<text:s/>b`), and the
 * qualified names ODF attributes are looked up by (`text:style-name`), as PHP's
 * DOM `getAttribute()` does.
 *
 * It is strict where DOM is strict about structure (an unclosed or mismatched
 * element is an error, so a broken part is refused rather than half-read), and
 * it never expands anything but the five predefined entities and character
 * references. A DOCTYPE is refused before this runs.
 */

export interface DomElement {
  /** Qualified name as written, e.g. `text:p`. */
  name: string;
  /** Local name, e.g. `p`. */
  local: string;
  /** Attributes by qualified name. */
  attrs: Map<string, string>;
  children: DomNode[];
}

export type DomNode = DomElement | string;

export function isElement(node: DomNode | undefined): node is DomElement {
  return typeof node === "object" && node !== null;
}

export function parseDom(src: string): DomElement {
  const n = src.length;
  const root: DomElement = { name: "#document", local: "#document", attrs: new Map(), children: [] };
  const stack: DomElement[] = [root];
  let i = 0;

  const text = (s: string): void => {
    if (s === "") return;
    const parent = stack[stack.length - 1]!;
    if (parent === root) {
      if (/[^ \t\r\n]/.test(s)) throw new Error("text outside the root element");
      return;
    }
    const last = parent.children[parent.children.length - 1];
    if (typeof last === "string") parent.children[parent.children.length - 1] = last + s;
    else parent.children.push(s);
  };

  while (i < n) {
    if (src.charCodeAt(i) !== 0x3c) {
      const next = src.indexOf("<", i);
      const stop = next < 0 ? n : next;
      text(unescape(src.slice(i, stop).replace(/\r\n?/g, "\n")));
      i = stop;
      continue;
    }
    if (src.startsWith("<!--", i)) {
      const end = src.indexOf("-->", i + 4);
      if (end < 0) throw new Error("unterminated comment");
      i = end + 3;
      continue;
    }
    if (src.startsWith("<![CDATA[", i)) {
      const end = src.indexOf("]]>", i + 9);
      if (end < 0) throw new Error("unterminated CDATA section");
      text(src.slice(i + 9, end).replace(/\r\n?/g, "\n"));
      i = end + 3;
      continue;
    }
    if (src.startsWith("<?", i)) {
      const end = src.indexOf("?>", i + 2);
      if (end < 0) throw new Error("unterminated processing instruction");
      i = end + 2;
      continue;
    }
    if (src.startsWith("<!", i)) throw new Error("markup declarations are not allowed");
    if (src.startsWith("</", i)) {
      const end = src.indexOf(">", i + 2);
      if (end < 0) throw new Error("unterminated end tag");
      const name = src.slice(i + 2, end).trim();
      const open = stack.pop();
      if (open === undefined || open === root || open.name !== name) throw new Error(`mismatched end tag ${name}`);
      i = end + 1;
      continue;
    }

    const end = tagEnd(src, i + 1);
    if (end < 0) throw new Error("unterminated start tag");
    let inner = src.slice(i + 1, end);
    const selfClosing = inner.endsWith("/");
    if (selfClosing) inner = inner.slice(0, -1);
    const element = startTag(inner);
    const parent = stack[stack.length - 1]!;
    if (parent === root && root.children.length > 0) throw new Error("more than one root element");
    parent.children.push(element);
    if (!selfClosing) stack.push(element);
    i = end + 1;
  }

  if (stack.length !== 1 || !isElement(root.children[0])) throw new Error("the document is incomplete");
  return root.children[0];
}

function tagEnd(src: string, from: number): number {
  let quote = 0;
  for (let i = from; i < src.length; i++) {
    const ch = src.charCodeAt(i);
    if (quote !== 0) {
      if (ch === quote) quote = 0;
    } else if (ch === 0x22 || ch === 0x27) {
      quote = ch;
    } else if (ch === 0x3e) {
      return i;
    }
  }
  return -1;
}

const ATTR = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function startTag(inner: string): DomElement {
  const m = /^\s*([^\s]+)/.exec(inner);
  if (!m) throw new Error("empty start tag");
  const name = m[1]!;
  const attrs = new Map<string, string>();
  for (const a of inner.slice(m[0].length).matchAll(ATTR)) {
    // Attribute-value normalisation: literal white space becomes a space.
    const raw = (a[2] ?? a[3] ?? "").replace(/\r\n?/g, "\n").replace(/[\t\n\r]/g, " ");
    if (!attrs.has(a[1]!)) attrs.set(a[1]!, unescape(raw));
  }
  const colon = name.indexOf(":");
  return { name, local: colon >= 0 ? name.slice(colon + 1) : name, attrs, children: [] };
}

function unescape(s: string): string {
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z]+);/g, (whole, ent: string) => {
    switch (ent) {
      case "amp": return "&";
      case "lt": return "<";
      case "gt": return ">";
      case "quot": return '"';
      case "apos": return "'";
    }
    if (ent[0] !== "#") throw new Error(`undefined entity ${whole}`);
    const code = ent[1] === "x" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
    if (!(code === 0x9 || code === 0xa || code === 0xd || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || (code >= 0x10000 && code <= 0x10ffff))) {
      throw new Error(`invalid character reference ${whole}`);
    }
    return String.fromCodePoint(code);
  });
}

/** All descendant text, as DOM `textContent`. */
export function textContent(node: DomNode): string {
  if (typeof node === "string") return node;
  return node.children.map(textContent).join("");
}
