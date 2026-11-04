import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parseDom } from "../src/reader/odt/xml-dom";
import { declaresDoctype, prologDeclaresDoctype } from "../src/reader/part-xml";
import { parseXml } from "../src/reader/xml";

const DOCTYPE = `<!DOCTYPE r [<!ENTITY x "LWINT">]>`;
const nul = (s: string): string => [...s].map((c) => c + "\0").join("");

describe("the prolog check", () => {
  it.each([
    ["bare", `${DOCTYPE}<r/>`],
    ["after a declaration", `<?xml version="1.0"?>${DOCTYPE}<r/>`],
    ["after whitespace and a comment", `<?xml version="1.0"?>\n <!-- c -->\n${DOCTYPE}<r/>`],
    ["after a BOM", `﻿${DOCTYPE}<r/>`],
    ["past a long comment", `<!-- ${"p".repeat(100000)} -->${DOCTYPE}<r/>`],
    ["after a processing instruction", `<?pi a?>${DOCTYPE}<r/>`],
    ["UTF-16 decoded as UTF-8, LE", `��${nul(`<?xml version="1.0"?>${DOCTYPE}<r/>`)}`],
    ["UTF-16 decoded as UTF-8, no BOM", nul(`<?xml version="1.0"?>${DOCTYPE}<r/>`)],
  ])("finds one: %s", (_name, xml) => {
    expect(prologDeclaresDoctype(xml)).toBe(true);
  });

  it.each([
    ["empty", ""],
    ["a root", "<r/>"],
    ["text after the root", "<r><![CDATA[<!DOCTYPE html>]]></r>"],
    ["in a comment", "<!-- <!DOCTYPE x> --><r/>"],
    ["in a processing instruction", "<?pi <!DOCTYPE x>?><r/>"],
    ["an unterminated comment", `<!-- ${DOCTYPE}`],
    ["an unterminated declaration", `<?xml ${DOCTYPE}`],
  ])("finds none: %s", (_name, xml) => {
    expect(prologDeclaresDoctype(xml)).toBe(false);
  });
});

describe("the gate", () => {
  it("is in both parsers", () => {
    expect(parseXml(`${DOCTYPE}<r/>`)).toBeNull();
    expect(() => parseDom(`${DOCTYPE}<r/>`)).toThrow();
    expect(parseXml(`<?xml version="1.0"?><r a="1"><b>t</b></r>`)?.children[0]?.text).toBe("t");
  });

  it("reads text that merely mentions a DOCTYPE", () => {
    expect(parseXml("<r><![CDATA[<!DOCTYPE html>]]></r>")?.text).toBe("<!DOCTYPE html>");
    expect(declaresDoctype("<r><![CDATA[<!DOCTYPE html>]]></r>")).toBe(false);
  });

  it("asks the parser as well: a DOCTYPE the walk cannot reach is still one", () => {
    // Text before the DOCTYPE ends the walk at once, and the parser still meets it.
    const hidden = `x${DOCTYPE}<r/>`;
    expect(prologDeclaresDoctype(hidden)).toBe(false);
    expect(parseXml(hidden)).toBeNull();
    expect(declaresDoctype(hidden)).toBe(true);
  });
});

/** "file: what" for each way past the gate in a TypeScript source tree. */
function bypasses(root: string): string[] {
  const found: string[] = [];
  const walk = (dir: string): void => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) {
        walk(path);
      } else if (name.endsWith(".ts") && name !== "part-xml.ts") {
        const src = readFileSync(path, "utf8");
        for (const [what, re] of [
          ["DOMParser", /\bnew\s+DOMParser\b/],
          ["XMLParser", /\bnew\s+XMLParser\b/],
          ["a byte scan", /<!DOCTYPE\/i|includes\("<!DOCTYPE/],
        ] as const) {
          if (re.test(src)) found.push(`${name}: ${what}`);
        }
        // Each parser's own entry point is the gate; nothing else may walk markup.
        if (/from "(?:\.\.?\/)*(?:[\w-]+\/)*(?:saxes|sax|fast-xml-parser|xmldom)/.test(src)) found.push(`${name}: a parser package`);
      }
    }
  };
  walk(root);
  return found;
}

describe("nothing else parses", () => {
  it("leaves no way to a part except through the gate", () => {
    expect(bypasses(join(__dirname, "..", "src"))).toEqual([]);
  });

  it("the entry points both start at the gate", () => {
    for (const [file, fn] of [["reader/xml.ts", "parseXml"], ["reader/odt/xml-dom.ts", "parseDom"]] as const) {
      const src = readFileSync(join(__dirname, "..", "src", file), "utf8");
      const at = src.indexOf(`export function ${fn}(`);
      expect(src.slice(at, at + 200), file).toContain("prologDeclaresDoctype(src)");
    }
  });
});
