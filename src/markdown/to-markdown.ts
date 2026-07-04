/**
 * Doc → GFM markdown — the Editor bridge. Hand-rolled (no external markdown
 * dependency), mirroring PHP `Markdown\MarkdownBridge::toMarkdown`.
 *
 * Lossy by design where GFM has no syntax: underline / color / highlight
 * decorations, paragraph alignment, image pixel sizes, and page breaks are
 * dropped. Everything else round-trips through `fromMarkdown`.
 */

import type { Block, Doc, ListItem, Run } from "../schema/types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

export function toMarkdown(doc: Doc): string {
  const chunks: string[] = [];
  for (const block of doc.blocks ?? []) {
    const md = blockToMarkdown(block as Any);
    if (md !== null) chunks.push(md);
  }
  return chunks.join("\n\n") + (chunks.length > 0 ? "\n" : "");
}

function blockToMarkdown(block: Any): string | null {
  switch (block?.type) {
    case "heading":
      return "#".repeat(clamp(block.level)) + " " + inline(block.runs ?? []);
    case "paragraph":
      return guardLineStart(inline(block.runs ?? []).replace(/\n/g, " "));
    case "list":
      return listToMarkdown(block.items ?? [], block.ordered === true, 0).join("\n");
    case "table":
      return tableToMarkdown(block);
    case "code": {
      const lang = typeof block.language === "string" ? block.language : "";
      return "```" + lang + "\n" + String(block.text ?? "") + "\n```";
    }
    case "quote": {
      const inner = (block.blocks ?? [])
        .map((b: Any) => blockToMarkdown(b))
        .filter((s: string | null): s is string => s !== null)
        .join("\n\n");
      return inner
        .split("\n")
        .map((line: string) => (line === "" ? ">" : "> " + line))
        .join("\n");
    }
    case "image": {
      const alt = typeof block.alt === "string" ? block.alt : "";
      return `![${alt.replace(/([[\]\\])/g, "\\$1")}](${String(block.src ?? "")})`;
    }
    case "hr":
      return "---";
    case "pageBreak":
      return null; // no GFM equivalent
    default:
      return null;
  }
}

function listToMarkdown(items: ListItem[], ordered: boolean, depth: number): string[] {
  const indent = (ordered ? "   " : "  ").repeat(depth);
  const lines: string[] = [];
  items.forEach((item, i) => {
    const marker = ordered ? `${i + 1}. ` : "- ";
    lines.push(indent + marker + inline(item.runs ?? []).replace(/\n/g, " "));
    if (item.children && item.children.length > 0) {
      lines.push(...listToMarkdown(item.children, ordered, depth + 1));
    }
  });
  return lines;
}

function tableToMarkdown(block: Any): string {
  const rows: Any[] = block.rows ?? [];
  if (rows.length === 0) return "";
  const cols = Math.max(1, ...rows.map((r: Any) => (Array.isArray(r?.cells) ? r.cells.length : 0)));

  const lineOf = (row: Any): string => {
    const cells: Any[] = Array.isArray(row?.cells) ? row.cells : [];
    const rendered: string[] = [];
    for (let c = 0; c < cols; c++) {
      rendered.push(cellInline(cells[c]));
    }
    return "| " + rendered.join(" | ") + " |";
  };

  const lines: string[] = [];
  lines.push(lineOf(rows[0]));
  lines.push("| " + Array.from({ length: cols }, () => "---").join(" | ") + " |");
  for (let r = 1; r < rows.length; r++) lines.push(lineOf(rows[r]));
  return lines.join("\n");
}

function cellInline(cell: Any): string {
  const blocks: Any[] = Array.isArray(cell?.blocks) ? cell.blocks : [];
  const parts: string[] = [];
  for (const b of blocks) {
    if (b?.type === "paragraph" || b?.type === "heading") parts.push(inline(b.runs ?? []));
    else if (b?.type === "code") parts.push("`" + String(b.text ?? "").replace(/\n/g, " ") + "`");
    else {
      const md = blockToMarkdown(b);
      if (md !== null) parts.push(md.replace(/\n/g, " "));
    }
  }
  return parts.join(" ").replace(/\|/g, "\\|").replace(/\n/g, " ");
}

// ── Inline runs → markdown ──────────────────────────────────────────────────

function inline(runs: Run[]): string {
  let out = "";
  for (const run of runs) {
    out += runToMarkdown(run);
  }
  return out;
}

function runToMarkdown(run: Run): string {
  if (typeof run?.text !== "string") return "";
  let s = run.code ? codeSpan(run.text) : escapeText(run.text);
  if (run.bold) s = `**${s}**`;
  if (run.italic) s = `*${s}*`;
  if (run.strike) s = `~~${s}~~`;
  if (run.link) s = `[${s}](${run.link})`;
  return s;
}

function codeSpan(text: string): string {
  if (!text.includes("`")) return "`" + text + "`";
  return "`` " + text + " ``";
}

function escapeText(text: string): string {
  return text.replace(/([\\`*_~[\]])/g, "\\$1");
}

/** Escape a leading character that would re-parse as block syntax. */
function guardLineStart(text: string): string {
  if (/^(#{1,6} |[-*+] |\d+\. |>|(-{3,}|\*{3,}|_{3,})$|\|)/.test(text)) {
    return "\\" + text;
  }
  return text;
}

function clamp(level: Any): number {
  const n = Number.isFinite(Number(level)) ? Math.trunc(Number(level)) : 1;
  return Math.min(6, Math.max(1, n));
}
