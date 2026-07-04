/**
 * Tiny isomorphic XML parser — enough to read OOXML parts (replacing PHP's
 * SimpleXML). Namespace prefixes are stripped so nodes/attributes are queried
 * by local name, matching how the PHP reader accesses SimpleXML.
 */

export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx >= 0 ? name.slice(idx + 1) : name;
}

function unescapeXml(s: string): string {
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|\w+);/g, (m, ent: string) => {
    switch (ent) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default:
        if (ent[0] === "#") {
          const code = ent[1] === "x" || ent[1] === "X" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
          return Number.isFinite(code) ? String.fromCodePoint(code) : m;
        }
        return m;
    }
  });
}

function findTagEnd(src: string, from: number): number {
  let quote = "";
  for (let i = from; i < src.length; i++) {
    const ch = src[i]!;
    if (quote) {
      if (ch === quote) quote = "";
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      return i;
    }
  }
  return src.length;
}

const ATTR_RE = /([^\s=/]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseTag(inner: string): { name: string; attrs: Record<string, string> } {
  const trimmed = inner.trim();
  let i = 0;
  while (i < trimmed.length && !/\s/.test(trimmed[i]!)) i++;
  const name = trimmed.slice(0, i);
  const attrs: Record<string, string> = {};
  const rest = trimmed.slice(i);
  for (const m of rest.matchAll(ATTR_RE)) {
    const key = localName(m[1]!);
    const value = unescapeXml(m[3] ?? m[4] ?? "");
    attrs[key] = value;
  }
  return { name, attrs };
}

export function parseXml(src: string): XmlNode | null {
  let i = 0;
  const n = src.length;
  const root: XmlNode = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XmlNode[] = [root];

  while (i < n) {
    if (src[i] === "<") {
      if (src.startsWith("<!--", i)) {
        const end = src.indexOf("-->", i);
        i = end < 0 ? n : end + 3;
        continue;
      }
      if (src.startsWith("<![CDATA[", i)) {
        const end = src.indexOf("]]>", i);
        stack[stack.length - 1]!.text += src.slice(i + 9, end < 0 ? n : end);
        i = end < 0 ? n : end + 3;
        continue;
      }
      if (src.startsWith("<?", i)) {
        const end = src.indexOf("?>", i);
        i = end < 0 ? n : end + 2;
        continue;
      }
      if (src.startsWith("<!", i)) {
        const end = src.indexOf(">", i);
        i = end < 0 ? n : end + 1;
        continue;
      }
      if (src[i + 1] === "/") {
        const end = src.indexOf(">", i);
        if (stack.length > 1) stack.pop();
        i = end < 0 ? n : end + 1;
        continue;
      }
      const end = findTagEnd(src, i + 1);
      const tagContent = src.slice(i + 1, end);
      const selfClosing = tagContent.endsWith("/");
      const { name, attrs } = parseTag(selfClosing ? tagContent.slice(0, -1) : tagContent);
      const node: XmlNode = { name: localName(name), attrs, children: [], text: "" };
      stack[stack.length - 1]!.children.push(node);
      if (!selfClosing) stack.push(node);
      i = end + 1;
    } else {
      const next = src.indexOf("<", i);
      const stop = next < 0 ? n : next;
      stack[stack.length - 1]!.text += unescapeXml(src.slice(i, stop));
      i = stop;
    }
  }

  return root.children[0] ?? null;
}

/** First child element with the given local name. */
export function el(node: XmlNode | null | undefined, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

/** All child elements with the given local name. */
export function els(node: XmlNode | null | undefined, name: string): XmlNode[] {
  return node ? node.children.filter((c) => c.name === name) : [];
}

/** Attribute by local name, or undefined. */
export function at(node: XmlNode | null | undefined, name: string): string | undefined {
  return node?.attrs[name];
}
