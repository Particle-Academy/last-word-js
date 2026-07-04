# @particle-academy/last-word

[![Fancy UI suite](art/fancy-ui.svg)](https://particle.academy)

Zero-dependency, **isomorphic** (browser + Node) `.docx` writer + reader for
agentic word-processing documents — a JSON document model with **markdown
bridges**. The Node/TypeScript mirror of the PHP
[`particle-academy/last-word`](https://github.com/Particle-Academy/last-word).
Sister to [`holy-sheet`](https://github.com/Particle-Academy/holy-sheet-js)
(xlsx) and [`dark-slide`](https://github.com/Particle-Academy/dark-slide-js)
(pptx).

The point is the **Editor round-trip**: a WYSIWYG editor (react-fancy
`Editor`) speaks markdown; Word speaks `.docx`. LastWord bridges the two
through one JSON model — `fromMarkdown → toBytes` to export a real Word file,
`read → toMarkdown` to import one — with no converter sandwich
(mammoth/turndown/docx) in between.

```ts
import { Agent } from "@particle-academy/last-word";

// Markdown in…
const doc = Agent.fromMarkdown(`# Q3 Report

Revenue was **up 12%** — see the [dashboard](https://example.com).

- Wins
  - Enterprise renewals
- Risks
`);

// …Word file out.
const bytes: Uint8Array = Agent.toBytes(doc); // universal
await Agent.write(doc, "report.docx"); // Node only

// And back: .docx → model → markdown for the editor.
const imported = Agent.read(bytes);
const markdown = Agent.toMarkdown(imported);
```

## The document model

A `Doc` is `{ title?, blocks }`. Blocks are JSON-friendly discriminated
unions — exactly what an agent emits:

| Block | Shape |
| --- | --- |
| heading | `{ type: "heading", level: 1-6, runs }` |
| paragraph | `{ type: "paragraph", runs, align? }` |
| list | `{ type: "list", ordered?, items: [{ runs, children? }] }` (nesting ≥ 3 deep) |
| table | `{ type: "table", rows: [{ header?, cells: [{ blocks }] }] }` |
| code | `{ type: "code", language?, text }` |
| quote | `{ type: "quote", blocks }` |
| image | `{ type: "image", src: "data:image/png;base64,…", widthPx?, heightPx?, alt? }` |
| pageBreak | `{ type: "pageBreak" }` |
| hr | `{ type: "hr" }` |

A `Run` is an inline span: `{ text, bold?, italic?, underline?, strike?,
code?, link?, color?, highlight? }` (colors are `#RRGGBB`).

## API

`Agent` (static) mirrors the PHP surface:

- `validate(doc)` → structured errors `{path, message}[]` (empty = valid)
- `validateAndRepair(doc)` → `{ok, schema, errors}` (coerces strings to runs,
  clamps heading levels, drops unknown block types with the error retained)
- `toBytes(doc)` → `Uint8Array` (universal, deterministic output)
- `write(doc, path)` → `{path, bytes, blocks}` (Node only)
- `read(bytes)` / `fromBytes(bytes)` → `Doc` (universal; tolerates
  Word-authored files — outlineLvl headings, named highlights, unknown
  constructs degrade to paragraphs, never throw)
- `toMarkdown(doc)` / `fromMarkdown(md)` → the Editor bridge (GFM: headings,
  `**`/`*`/`~~`, inline code, links, nested lists, tables, fenced code,
  blockquotes, images, `---`)
- `describe(doc)` → plain-text summary (title, block counts, word count)
- `jsonSchema()` → JSON Schema for LLM tool-use
- `version()` → package version

Markdown is lossy only where GFM has no syntax: underline / color / highlight
decorations, paragraph alignment, image pixel sizes, and page breaks are
dropped on `toMarkdown`; everything else round-trips.

Images are embedded from data URLs (PNG/JPEG); when `widthPx`/`heightPx` are
omitted the intrinsic size is sniffed from the bytes (PNG IHDR / JPEG SOF)
and capped at 6.5in width keeping aspect.

## Cross-language parity

As of 0.2.0 the metadata slots match the PHP mirror exactly: the title is
carried in `docProps/core.xml` (`dc:title`) and the code block `language` in
a `lastword:code:{lang}` content-control tag (quotes use `lastword:quote`),
so the **same file opens in either engine** — title and code language
round-trip Node ↔ PHP in both directions. Files written by PHP 0.1.x
(Title-styled paragraph, `LastWordCode_{lang}` bookmark) still read fine;
the sibling repo's canonical fixture is frozen into each test suite as a
cross-read vector.

---

## ⭐ Star Fancy UI

If this package is useful to you, a quick ⭐ on the repo really helps us build a better kit. Thank you!
