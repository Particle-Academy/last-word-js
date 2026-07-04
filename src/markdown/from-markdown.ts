/**
 * GFM markdown → Doc — the Editor bridge. Hand-rolled line-based block parser
 * plus an inline tokenizer (bold/italic/strike/code/links with backslash
 * escapes). No external markdown dependency. Mirrors PHP
 * `Markdown\MarkdownBridge::fromMarkdown`.
 */

import { mergeRuns } from "../reader/docx-reader";
import type { Block, Doc, ListItem, Run } from "../schema/types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const HEADING_RE = /^(#{1,6})\s+(.*)$/;
const HR_RE = /^(-{3,}|\*{3,}|_{3,})\s*$/;
const FENCE_RE = /^```(.*)$/;
const LIST_RE = /^(\s*)([-*+]|\d{1,9}\.)\s+(.*)$/;
const IMAGE_LINE_RE = /^!\[((?:\\.|[^\]\\])*)\]\(([^)]*)\)\s*$/;
const TABLE_SEP_RE = /^\s*\|?(\s*:?-{3,}:?\s*\|)*\s*:?-{3,}:?\s*\|?\s*$/;

export function fromMarkdown(markdown: string): Doc {
  const lines = String(markdown ?? "")
    .replace(/\r\n?/g, "\n")
    .split("\n");
  return { blocks: parseBlocks(lines) };
}

function parseBlocks(lines: string[]): Block[] {
  const blocks: Block[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i]!;

    if (line.trim() === "") {
      i++;
      continue;
    }

    // Fenced code block
    const fence = FENCE_RE.exec(line);
    if (fence) {
      const language = fence[1]!.trim();
      const body: string[] = [];
      i++;
      while (i < lines.length && !/^```\s*$/.test(lines[i]!)) {
        body.push(lines[i]!);
        i++;
      }
      i++; // closing fence (or EOF)
      const block: Any = { type: "code", text: body.join("\n") };
      if (language !== "") block.language = language;
      blocks.push(reorderCode(block));
      continue;
    }

    // Heading
    const heading = HEADING_RE.exec(line);
    if (heading) {
      blocks.push({
        type: "heading",
        level: heading[1]!.length as 1 | 2 | 3 | 4 | 5 | 6,
        runs: parseInline(heading[2]!),
      });
      i++;
      continue;
    }

    // Horizontal rule
    if (HR_RE.test(line)) {
      blocks.push({ type: "hr" });
      i++;
      continue;
    }

    // Blockquote
    if (line.startsWith(">")) {
      const inner: string[] = [];
      while (i < lines.length && lines[i]!.startsWith(">")) {
        inner.push(lines[i]!.replace(/^> ?/, ""));
        i++;
      }
      blocks.push({ type: "quote", blocks: parseBlocks(inner) });
      continue;
    }

    // Table
    if (line.trimStart().startsWith("|") && i + 1 < lines.length && TABLE_SEP_RE.test(lines[i + 1]!)) {
      const rows: Any[] = [];
      const headerCells = splitTableRow(line);
      rows.push({ header: true, cells: headerCells });
      i += 2; // header + separator
      while (i < lines.length && lines[i]!.trimStart().startsWith("|")) {
        rows.push({ cells: splitTableRow(lines[i]!) });
        i++;
      }
      blocks.push({ type: "table", rows } as Block);
      continue;
    }

    // List
    if (LIST_RE.test(line)) {
      const entries: { indent: number; ordered: boolean; content: string }[] = [];
      while (i < lines.length) {
        const m = LIST_RE.exec(lines[i]!);
        if (!m) break;
        entries.push({
          indent: m[1]!.length,
          ordered: /\d/.test(m[2]![0]!),
          content: m[3]!,
        });
        i++;
      }
      blocks.push(buildListBlock(entries));
      continue;
    }

    // Standalone image
    const image = IMAGE_LINE_RE.exec(line);
    if (image) {
      const block: Any = { type: "image", src: image[2]! };
      const alt = image[1]!.replace(/\\([[\]\\])/g, "$1");
      if (alt !== "") block.alt = alt;
      blocks.push(block as Block);
      i++;
      continue;
    }

    // Paragraph: consecutive plain lines soft-wrapped with a space.
    const para: string[] = [line];
    i++;
    while (i < lines.length && lines[i]!.trim() !== "" && !isBlockStart(lines[i]!)) {
      para.push(lines[i]!);
      i++;
    }
    blocks.push({ type: "paragraph", runs: parseInline(para.join(" ")) });
  }

  return blocks;
}

function isBlockStart(line: string): boolean {
  return (
    HEADING_RE.test(line) ||
    HR_RE.test(line) ||
    FENCE_RE.test(line) ||
    LIST_RE.test(line) ||
    IMAGE_LINE_RE.test(line) ||
    line.startsWith(">") ||
    line.trimStart().startsWith("|")
  );
}

// ── Lists ───────────────────────────────────────────────────────────────────

function buildListBlock(entries: { indent: number; ordered: boolean; content: string }[]): Block {
  const root: ListItem[] = [];
  const stack: { indent: number; items: ListItem[]; last: ListItem | null }[] = [
    { indent: entries[0]?.indent ?? 0, items: root, last: null },
  ];

  for (const entry of entries) {
    while (stack.length > 1 && entry.indent < stack[stack.length - 1]!.indent) {
      stack.pop();
    }
    let top = stack[stack.length - 1]!;
    if (entry.indent > top.indent && top.last) {
      if (!top.last.children) top.last.children = [];
      stack.push({ indent: entry.indent, items: top.last.children, last: null });
      top = stack[stack.length - 1]!;
    }
    const item: ListItem = { runs: parseInline(entry.content) };
    top.items.push(item);
    top.last = item;
  }

  const ordered = entries[0]?.ordered === true;
  return (ordered ? { type: "list", ordered: true, items: root } : { type: "list", items: root }) as Block;
}

// ── Tables ──────────────────────────────────────────────────────────────────

function splitTableRow(line: string): Any[] {
  let s = line.trim();
  if (s.startsWith("|")) s = s.slice(1);
  if (s.endsWith("|") && !s.endsWith("\\|")) s = s.slice(0, -1);

  const cells: string[] = [];
  let buf = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === "\\" && i + 1 < s.length) {
      buf += ch + s[i + 1]!;
      i++;
      continue;
    }
    if (ch === "|") {
      cells.push(buf);
      buf = "";
      continue;
    }
    buf += ch;
  }
  cells.push(buf);

  return cells.map((cell) => ({ blocks: [{ type: "paragraph", runs: parseInline(cell.trim()) }] }));
}

// ── Inline tokenizer ────────────────────────────────────────────────────────

const PUNCT_RE = /[!-/:-@[-`{-~]/;
const WORD_RE = /[A-Za-z0-9_]/;

export function parseInline(text: string): Run[] {
  const runs: Run[] = [];
  let buf = "";
  let bold = false;
  let italic = false;
  let strike = false;
  let i = 0;
  const len = text.length;

  const flush = (): void => {
    if (buf !== "") {
      runs.push(makeRun(buf, bold, italic, strike, false, undefined));
      buf = "";
    }
  };

  while (i < len) {
    const ch = text[i]!;
    const two = text.slice(i, i + 2);

    // Backslash escape
    if (ch === "\\" && i + 1 < len && PUNCT_RE.test(text[i + 1]!)) {
      buf += text[i + 1]!;
      i += 2;
      continue;
    }

    // Code span (single or multi-backtick fence)
    if (ch === "`") {
      const fence = /^`+/.exec(text.slice(i))![0];
      const end = text.indexOf(fence, i + fence.length);
      if (end >= 0) {
        flush();
        let content = text.slice(i + fence.length, end);
        if (fence.length > 1) content = content.replace(/^ /, "").replace(/ $/, "");
        runs.push(makeRun(content, bold, italic, strike, true, undefined));
        i = end + fence.length;
        continue;
      }
      buf += ch;
      i++;
      continue;
    }

    // Bold
    if (two === "**" || two === "__") {
      flush();
      bold = !bold;
      i += 2;
      continue;
    }

    // Strikethrough
    if (two === "~~") {
      flush();
      strike = !strike;
      i += 2;
      continue;
    }

    // Italic via *
    if (ch === "*") {
      flush();
      italic = !italic;
      i++;
      continue;
    }

    // Italic via _ (word-boundary aware so snake_case survives)
    if (ch === "_") {
      const prev = i > 0 ? text[i - 1]! : " ";
      const next = i + 1 < len ? text[i + 1]! : " ";
      if ((!italic && !WORD_RE.test(prev)) || (italic && !WORD_RE.test(next))) {
        flush();
        italic = !italic;
        i++;
        continue;
      }
      buf += ch;
      i++;
      continue;
    }

    // Link
    if (ch === "[") {
      const match = matchLink(text, i);
      if (match) {
        flush();
        const inner = parseInline(match.label);
        for (const run of inner) {
          run.link = match.url;
          if (bold) run.bold = true;
          if (italic) run.italic = true;
          if (strike) run.strike = true;
        }
        runs.push(...inner);
        i = match.end;
        continue;
      }
      buf += ch;
      i++;
      continue;
    }

    // Inline image: degrade to its alt text (the model has no inline images).
    if (ch === "!" && text[i + 1] === "[") {
      const match = matchLink(text, i + 1);
      if (match) {
        buf += match.label.replace(/\\([[\]\\])/g, "$1");
        i = match.end;
        continue;
      }
    }

    buf += ch;
    i++;
  }
  flush();

  return mergeRuns(runs);
}

function makeRun(
  text: string,
  bold: boolean,
  italic: boolean,
  strike: boolean,
  code: boolean,
  link: string | undefined,
): Run {
  const run: Run = { text };
  if (bold) run.bold = true;
  if (italic) run.italic = true;
  if (strike) run.strike = true;
  if (code) run.code = true;
  if (link) run.link = link;
  return run;
}

function matchLink(text: string, start: number): { label: string; url: string; end: number } | null {
  let depth = 0;
  let labelEnd = -1;
  for (let j = start; j < text.length; j++) {
    const ch = text[j]!;
    if (ch === "\\") {
      j++;
      continue;
    }
    if (ch === "[") depth++;
    else if (ch === "]") {
      depth--;
      if (depth === 0) {
        labelEnd = j;
        break;
      }
    }
  }
  if (labelEnd < 0 || text[labelEnd + 1] !== "(") return null;

  for (let k = labelEnd + 2; k < text.length; k++) {
    const ch = text[k]!;
    if (ch === "\\") {
      k++;
      continue;
    }
    if (ch === ")") {
      return { label: text.slice(start + 1, labelEnd), url: text.slice(labelEnd + 2, k), end: k + 1 };
    }
  }
  return null;
}

/** Keep `language` before `text` for stable JSON output (cosmetic only). */
function reorderCode(block: Any): Block {
  if (block.language === undefined) return block as Block;
  return { type: "code", language: block.language, text: block.text } as Block;
}
