# Changelog

All notable changes to `@particle-academy/last-word` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed

- **`CHANGELOG.md` is now in the published tarball.** `files` did not whitelist it, so npm never shipped it — and this package puts breaking changes in MINOR releases and tells you in the README to read the entry before taking one. The instruction existed for the author, who has the file, and not for the consumer, who is the only one being instructed. Nothing for you to do; the file simply arrives from this release on.

### Security

- `source-map-js` is pinned forward to `^1.2.2` via `overrides`. Versions up to
  1.2.1 allow an event-loop denial of service through indexed source-map section
  offsets, and it arrives here transitively through the build toolchain.
  **Nothing for a consumer to do, and no runtime change**: an npm package does
  not ship a lockfile, so this governs builds OF this repo, not anything
  installed FROM it. Recorded rather than left silent because the override it
  sits beside — `shell-quote` `^1.9.0`, added for an earlier advisory — was
  carried with no note of why, and had drifted back inside the vulnerable range
  before anyone looked.

## 0.6.1 — 2026-09-14

### Fixed

Mirrors `particle-academy/last-word` 0.6.3, whose fixes this port's 0.6.0
reported. Each has a test that fails against 0.6.0, and the PHP parity suite
runs against 0.6.3.

- **`diff()` emits every `doc.set` `key` as a string.** 0.6.0 matched PHP 0.6.2,
  which stored a top-level key `"5"` as the int 5 and emitted `key: 5`; the
  reducer refuses a non-string key, so the replay check failed and the whole
  diff became one `doc.replace`. It is now a `doc.set` with `key: "5"`, as PHP
  0.6.3 emits.
- **An op whose path names a non-empty list by a key is skipped.** An insert at
  `/blocks/blocks` turned the top-level block list into an object holding the
  old blocks under `"0"`, `"1"`, … A list is reached by position only. An empty
  list still takes a name, as in PHP, which cannot tell `[]` from `{}`.
- **`opSchema()` requires a non-empty `doc.set` `key`** (`minLength: 1`), as the
  reducer does. The schema is still byte-identical to PHP's.

  **What you must do:** nothing. A stored op list replays as before unless it
  held one of the paths above, which no `diff()` emits; and a history that fell
  back to `doc.replace` for a numeric key still replays correctly.

## 0.6.0 — 2026-09-14

### Added

- **`Agent.diff()`, `Agent.reduce()`, `Agent.opSchema()` and `Agent.equivalent()`:
  a document's versions stored as ops** (last-word#2). Ported from
  `particle-academy/last-word` 0.6.2, which is the reference (its 0.6.0 design
  with the 0.6.1 and 0.6.2 fixes): the same algorithm, so the same two documents
  give the same ops in the same order in both runtimes, and an op history
  written by one replays in the other. A version history cannot keep a .docx
  per edit, hashing the bytes cannot keep a one-word edit small (it is a zip),
  and diffing `toMarkdown()` output would lose run formatting, tables and page
  breaks on restore. The diff is over Last Word's own model.
  - `reduce(a, diff(a, b))` equals `b`, key order aside. The ops are verified
    by replaying them; ops that do not reproduce `b` become one `doc.replace`.
  - Every list is aligned by content: the top-level blocks, a quote's blocks, a
    list's items and their children, a table's rows, a row's cells and a cell's
    blocks. Rewording one paragraph is one `blocks.replace` at its own path, even
    inside a table cell; moving one is one `blocks.move`. A container whose own
    properties changed is replaced whole.
  - Documents that write the same file diff to `[]`, so a save without a change
    records nothing, even where the reader normalises (merged runs, a header
    row's bold, a dropped empty paragraph).
  - Blocks have no ids, so ops address a list by JSON Pointer and an item by
    index: `blocks.*`, `items.*`, `rows.*` and `cells.*`, each with
    `insert`/`remove`/`move`/`replace`, plus `doc.set` and `doc.replace`.
  - `reduce()` skips an op whose position is not an integer or a digit string,
    or whose `op`, `path` or `doc.set` `key` is not a string, rather than reading
    junk as position 0 (PHP 0.6.2).
  - The equality `diff()` runs on throws a `TypeError` for a value JSON cannot
    hold — a non-finite number, a string with a lone surrogate, a function,
    symbol or bigint, or more than 4096 nested arrays and objects — rather than
    calling two different ones the same (PHP 0.6.1 throws `JsonException`).

  **What you must do:** nothing. This only adds methods.

- **`DocOp`**, a discriminated union on `op` (with a type per op), plus
  **`DocDiff`**, **`DocReducer`** and **`DocOpSchema`**, exported beside the
  other building blocks.

`tests/doc-ops-parity.test.ts` runs the PHP package's `Agent::diff` and
`Agent::reduce` on every case and requires the same ops, field for field and in
order, and the same replayed document: each edit PHP's own suite pins, both
ways; seeded random edits at every depth; one entry moved in each kind of list;
raw alignments, where ties decide the hunks; and reducer edge cases no diff
emits. The op schema is compared byte for byte. Two differences cannot be
removed, because JS values do not carry the distinction: PHP calls the JSON
number `1.0` different from `1` (and skips `2.0` as a position), where this port
sees one number; and PHP compares a map keyed 0..10 or more out of order as a
map, where a JS object is always in key order. No valid document hits either.

## 0.5.0 — 2026-09-13

### Added

- **`Agent.read()` reads legacy Word `.doc` (Word 97-2003), `.odt` and `.rtf`,
  not just `.docx`** (last-word#1), matching the PHP engine. The format is
  decided from the bytes, never a file name, and all four return the same
  document shape. The compound-file (MS-CFB) and Word binary (MS-DOC) readers
  are this package's own code; there is still no dependency.

  - `.doc`: paragraph text from every piece (8-bit and UTF-16), headings by
    built-in style id, direct bold / italic / underline / strike, `HYPERLINK`
    fields (other fields keep their result), nested bulleted and numbered
    lists, tables with header rows, page breaks. Not read: style-inherited
    formatting, fonts / sizes / colours, images, text boxes, headers / footers
    / footnotes / comments, merged cells, the title.
  - `.odt`: headings, paragraphs, bold / italic / underline / strike from
    automatic styles, links, nested lists, tables with header rows and merged
    cells, `text:s` / `text:tab` / `text:line-break`, page breaks, the title.
  - `.rtf`: a group-scoped tokenizer; headings by style name or outline level,
    direct formatting with style formatting subtracted, links, lists from the
    list table, tables (header rows where `\trhdr` marks them), `\uN` with
    `\ucN` skipping and surrogate pairs, `\'hh` in the `\ansicpg` code page
    (874 and 1250-1258 decoded; 932/936/949/950 double-byte characters become
    one U+FFFD each).

  One document written by last-word and converted by LibreOffice reads back
  from `.doc` and `.odt` IDENTICAL to the `.docx`, and from `.rtf` identical
  except the header-row flag that file does not carry.
  `legacy-formats.test.ts` asserts it against `report.read.json`, the same
  committed answer the PHP and Python suites assert, compared as JSON text.

- **`UnsupportedFormatException`**, with a `format` naming what the bytes are:
  `doc` (Word 6/95 or encrypted), `xls` / `ppt` / `msg` / `cfb` (another
  compound file), `xlsx` / `pptx` / `ods` / `odp`, or `unknown`. A damaged
  file in a supported format throws a plain `Error`, because "this file is
  broken" and "save it as .docx" send a person to do different things.

- `DocReader`, `OdtReader`, `RtfReader`, `Format` and `detectFormat` are
  exported beside `DocxReader`.

### Changed

- **`Agent.read()` of bytes that are not a docx throws
  `UnsupportedFormatException`** where it used to fail inside the zip reader
  with whatever error that produced. It extends `Error`, so a `catch` still
  catches it; do nothing unless you matched the old message.

### Fixed

- **A truncated DEFLATE stream is an error, not garbage.** Inflating read past
  the end of its input as zero bits and returned whatever those decoded to, so
  a cut-off `.docx` part could come back as nonsense instead of failing.

### Security

- **Legacy readers treat an upload as hostile.** Compound-file offsets are
  bounds-checked; sector chains, the DIFAT chain and the directory tree are
  followed at most once per node; a stream over 256 MB, or an allocation table
  naming more sectors than the file holds, is refused; the Word piece table
  must run forwards; an ODT part carrying a DOCTYPE is refused; only the three
  ODT parts read are inflated, each refused past 64 MB declared or actual (a
  zip bomb stops at the cap); ODT element nesting past 257 is refused, the
  limit libxml puts on the PHP engine; repeated rows and columns are capped at 1,000
  per repeat and 100,000 cells in total; RTF nesting is capped at 10,000 and
  `\bin` data is skipped by its length. Each guard has a test on a hand-built
  file (`tests/support/legacy-files.ts`).

## 0.4.0 — 2026-09-10

### Added

- **A premium-document composition suite.** `PremiumDocumentTest` /
  `premium-document.test.ts` writes one document using every formatting feature
  at once — all eight run properties in a single paragraph, styled headings, a
  repeating table header, resolved list numbering, A4 landscape geometry, a
  document default font — and asserts the unzipped OOXML rather than that a file
  appeared.

  The suites beside it prove each feature works ON ITS OWN. That is a different
  question from whether they compose, and a dropped feature is invisible: Word
  shows no error and the document is merely plain. Nobody files a bug against a
  report that looks boring; they conclude the library is boring.

  It found no defects here, which is the result worth recording. It includes a
  control that fails if the assertions could pass on an unformatted document,
  and it pins two decisions that read like gaps: `highlight` renders as `<w:shd>`
  rather than `<w:highlight>` (which takes sixteen named colours and could not
  carry a `#RRGGBB` schema), and `page.margins` are in POINTS.

- **A rich-layout surface, so a business one-pager is expressible.** The model
  was far narrower than the XML this engine already emitted: font size, font
  family, small caps, letter spacing, per-cell shading, borders, padding,
  vertical alignment and both merge directions were produced from hardcoded
  blocks or from `styles.xml` and were **unreachable from the model**. An agent
  could emit `size`, `colSpan` or `shading`, the validator returned no errors,
  and every one of them was silently dropped.

  | where | new keys |
  |---|---|
  | run | `size` (points, half-points exact) · `font` · `smallCaps` · `letterSpacing` (points, may be negative) |
  | paragraph, **heading** and list item | `spaceBefore` · `spaceAfter` · `lineHeight` · `indentLeft` · `indentRight` · `keepNext` · `shading` · `borders` · `align` (on headings too) |
  | table | `widths` (relative column weights) · `width` (% of the text column) · `align` · `borders` (incl. `insideH` / `insideV`) · `cellPadding` |
  | cell | `shading` · `borders` · `padding` · `valign` · `colSpan` · `rowSpan` |
  | document | `page` (`size`, `orientation`, `margins`) · `defaultFont` · `defaultSize` |

  Every key is validated, every key round-trips through the reader, and every
  key appears in `jsonSchema()` so an agent registering the tool is told it
  exists.

- **A heading is now a paragraph.** It takes the same properties, so a section
  label that needs spacing or alignment no longer has to be a bold paragraph
  impersonating a heading — and therefore appearing in no navigation pane and
  no table of contents.

- **Both merge directions**, written HTML-style: a `rowSpan` cell appears ONCE
  and the rows it covers list only their own remaining cells. The writer
  synthesises the `w:vMerge` continuations OOXML requires and the reader folds
  them back, so the model that comes out is the model that went in.

- **The table grid is computed from the section.** All three engines carried
  `9360` twips as a literal, so a document that narrowed its margins got a
  table that no longer matched its own page — too narrow, and silently so.

- **`last-word/docx-constructs` in `fancy-conformance`** — 44 shared rows
  pinning which construct emits which XML, in which order, and what the reader
  gives back. The rows are not transcribed into this repo: all three engines
  assert the same file, so a mapping that drifts in one fails there rather than
  quietly becoming that engine's behaviour.

### Fixed
- **`version()` reports the version this package actually ships as.** It
  returned `0.2.0` from a 0.4.0 release. The constant had drifted because
  nothing compared it to the packaging metadata — the same shape as every other
  two-copies-of-one-number failure in this estate.

  `VersionIsSingleSourcedTest` / `version.test.ts` now pins it, so the class is
  closed rather than the instance fixed. `dark-slide-py` already had that
  assertion and was the only engine in the family to catch itself.


- **Adjacent tables no longer merge into one in Word.** OOXML merges two
  `<w:tbl>` elements that touch, imposing the first table's column grid on the
  second. A stat band followed by a callout became a single two-row table.
- **A run's properties can no longer be lost to run-merging.** Adjacent runs
  are merged when their formatting matches, and the comparison listed only the
  properties that existed when it was written — so two differently-sized runs
  merged into one and took the first one's size.

- **The `LastWordTable` style is no longer emitted.** Nothing references it
  now that table properties are inline. If you were overriding it in a template,
  override the table's `borders` and `cellPadding` instead.

### Changed

- **Table properties are now written inline, not taken from a named style.**
  A named table style cannot vary per table instance, so per-table borders
  forced this. Also reconciled with it: header cells are one grey in all three
  engines, and header bold is one mechanism.

- **Document defaults name an East Asian font.** Without it Word picks its own
  face for CJK runs, which is exactly the text a mixed-script document
  contains.

  **What you must do: nothing.** No existing key changed meaning, nothing was
  removed and nothing was renamed. A document written before this release
  produces the same page. The visible differences are confined to tables, are
  small, and are listed above so a pixel comparison against an old build is not
  a surprise.

### Notes

- **A `header: true` row is not a round-trip fixpoint, and now says so.** The
  writer bolds the row's runs and the reader honestly reports the bold it
  finds, so the model that comes out is not the model that went in. The
  alternatives were to stop bolding header rows (changing every existing
  consumer's output) or to have the reader strip bold from header rows
  (discarding bold an author really asked for). Pinned as case `0042` rather
  than left as a surprise.

- **`fancy-conformance` is a dev dependency only.** The shared table is on the
  registry at `0.20.0` and this package requires `^0.20.0`, so nothing here is
  gated on it. (An earlier draft of this entry said the release was BLOCKED on
  `fancy-conformance` 0.7.0 reaching a registry. That was true when it was
  written and stopped being true thirteen minors ago — the note outlived the
  condition, which is the failure mode of writing a blocker down and not dating
  it.) Nothing about the runtime surface depends on it; only the test that
  asserts the shared rows.

- **What DOCX cannot do, so nobody chases it:** table corners are always
  square (there is no border radius in WordprocessingML), a background cannot
  bleed past the page margin without an anchored drawing, and naming a font is
  not shipping one — a reader without it substitutes. None of the three is
  worked around here; a layout that needs them needs a different format.

### Added

- **Cross-engine WRITER parity against the PHP `last-word`** — the guarantee
  this pair was missing. `cross-read.test.ts` already proved the Node *reader*
  restores a frozen docx PHP wrote; nothing compared the two *writers*. Both
  siblings in this family (`holy-sheet`, `dark-slide`) diff PHP output against
  Node's; this one did not, so the engines could drift apart on anything the
  frozen fixture happened not to contain.

  Five documents — minimal, every inline mark, structure (lists / table /
  quote), non-ASCII (CJK, emoji, combining marks), and empty-vs-zero values —
  written by both engines and diffed part by part. Containers are never
  byte-compared: PHP writes DEFLATE with real mtimes, this port writes STORE
  with a fixed 1980 date, so the files can never match and it would be wrong to
  try.

  **PHP is REQUIRED in CI**, and the suite throws rather than skipping when it
  is absent. A skip is a green build with zero parity coverage, which is how the
  other two suites in this family reported success over nothing for months.

### Notes

- **The pair is NOT at parity, and the first run of the new suite proved it.**
  Four parts differ, and they drift in BOTH directions — so this is not "the
  port is behind":

  | part | difference |
  |---|---|
  | `word/document.xml` | hyperlinks: port emits `w:history="1"`, PHP does not · lists: PHP emits `<w:pStyle w:val="ListParagraph"/>`, port does not · tables: port references `<w:tblStyle w:val="LastWordTable"/>`, PHP inlines `<w:tblBorders>` |
  | `word/styles.xml` | port emits `w:eastAsia="Calibri"` |
  | `word/numbering.xml` | port emits `<w:multiLevelType w:val="hybridMultilevel"/>` |
  | `[Content_Types].xml` | PHP declares png/jpeg defaults unconditionally |

  **The table difference is user-visible**: borders from a named style versus
  borders written into the table properties. A table written by one backend is
  not the same object as one written by the other.

  Recorded in `KNOWN_DIVERGENT_PARTS` rather than fixed, because reconciling
  them changes rendered output for every existing consumer, and in several cases
  the port is arguably the more correct side — so "make the port match PHP", the
  obvious move given PHP is the reference, would remove the better behaviour.
  That is the pair owner's decision, not a side effect of writing the suite that
  found it.

  The list ratchets both ways: a new divergent part fails, and an entry that
  stops being true fails too. `word/document.xml` for the *minimal* document is
  asserted identical unconditionally — whatever else drifts, the engines must
  agree on a single plain paragraph.

## 0.3.0 — 2026-08-07

### Changed

- **BREAKING — Node 18 is no longer supported.** `engines.node` moves from `>=18` to `>=22`.

  **What you must do:** on Node 22 or newer, nothing. Note npm only *warns* on an `engines` mismatch while **pnpm fails the install**, so this surfaces differently depending on your package manager. Node 18 is end-of-life and 20 is maintenance-only.

### Why

These are the kit 0.5 platform floors, applied across every package at once so a consumer never has to resolve a mix. **No API changed, nothing was removed, nothing was renamed** — only what the package requires.

## 0.2.0

Cross-language metadata parity with the PHP mirror
(particle-academy/last-word) — last-word-js#1.

- The two metadata slots that didn't cross languages now do, both
  directions: **title** (docProps/core.xml `dc:title`) and **code block
  `language`** (`lastword:code:{lang}` w:sdt tag — canonical on both
  sides now).
- Reader: back-compat fallback for the PHP ≤0.1.x legacy slot — an
  invisible `LastWordCode_{lang}` bookmark on the first code paragraph.
- New frozen cross-read vector: `test/fixtures/php-canonical.docx`
  (written by the PHP engine) + its JSON, asserted semantically
  deep-equal on read.

## 0.1.0

Initial release — the docx sibling of holy-sheet (xlsx) and dark-slide (pptx).

- JSON document model: heading / paragraph / list (nested) / table / code /
  quote / image / pageBreak / hr blocks with styled runs (bold, italic,
  underline, strike, inline code, link, color, highlight).
- `Agent` façade: `validate`, `validateAndRepair`, `toBytes`, `write`,
  `read`/`fromBytes`, `toMarkdown`, `fromMarkdown`, `describe`, `jsonSchema`,
  `version` — mirrors the PHP `particle-academy/last-word` surface.
- DOCX writer: deterministic OOXML output (styles, numbering, hyperlink +
  image rels, media parts, EMU extents with PNG IHDR / JPEG SOF sniffing).
- DOCX reader: round-trips its own output and tolerates Word-authored files
  (outlineLvl headings, named highlights, unknown numIds, unknown elements
  degrade to paragraphs).
- Markdown bridges: hand-rolled GFM emitter + parser (no external markdown
  dependency) for the react-fancy Editor round-trip.
