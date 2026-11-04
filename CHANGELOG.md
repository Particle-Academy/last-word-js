# Changelog

All notable changes to `@particle-academy/last-word` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.14.0] - 2026-11-03

**What you must do: only if your documents carry `pageBreak.orientation`,
`pageBreak.margins` (or any page key straight on a break) or a top-level
`sections`.** Those were silently ignored and are now errors. Remove them, or
use `pageBreak.page`, which `fine-print` honours for PDFs today and `last-word`
will honour in a later release. Until then `last-word` refuses it. A document
without those keys writes exactly what 0.13.0 wrote, byte for byte. Mirrors PHP
`particle-academy/last-word` [Unreleased], with the same messages.

### Changed

- **BREAKING: `pageBreak.orientation`, `pageBreak.margins`, `.size`, `.header`,
  `.footer`, `.firstHeader`, `.firstFooter`, `.evenHeader`, `.evenFooter`,
  `.background`, `.border` and a top-level `sections` were silently ignored and
  are now errors.** The document came out all portrait and looked fine, with no
  word that the request had been lost. The message names the form to use:
  `a pageBreak has no orientation key; use pageBreak.page`, and
  `there is no sections key; use pageBreak.page`. Every other extra key on a break
  stays lenient. **Consumer action:** remove those keys, or for mixed pages use
  `pageBreak.page`, which is not honoured here until a later release (below).

### Added

- **`pageBreak.page` is refused until it can be written** (mixed page orientation,
  stage 1): `pageBreak.page is not implemented in last-word yet; remove page to
  keep one page setup, or write this document as a PDF with fine-print`. It applies
  at any depth, to `validate`, `toBytes` and `write`. `validateAndRepair` keeps the
  key and returns `ok: false`.
- **Reading a file with section breaks now says what it flattened.** `read()` has
  always taken one page setup from such a file, the last section's, silently. It
  now adds `unrepresented: [{where: "/page", kind: "layout-not-preserved", count}]`,
  `count` being the section breaks lost, and `describe()` gains a line for it. Not
  breaking: blocks, `page` and everything else read as before, and the key appears
  only on such a file. `Doc` gains the optional `unrepresented` and the new
  `UnrepresentedNote` type. This is the first use of `unrepresented` in this
  package; PHP also reports what a header or footer could not hold, which this
  reader does not yet.

## [0.13.0] - 2026-11-01

**What you must do: nothing, unless you want first or even bands, the new
header/footer fields or `meta.author`.** With no variant band, no `meta` and
none of the new tokens, every document writes exactly what 0.12.0 wrote, part
for part. Two read-side changes and one write-side change can touch an existing
document; see Changed. Mirrors PHP `particle-academy/last-word` 0.16.0, less the
letterhead option (see Notes). Pre-1.0: a minor release may carry breaking
changes, and this one carries none that need action.

### Added

- **First-page and even-page bands: `page.firstHeader`, `page.firstFooter`,
  `page.evenHeader`, `page.evenFooter`** (MOIC, 2026-11-01), each the same
  shape as `page.header`. `page.header` stays "every page"; with a
  `firstHeader` it is every page but the first, with an `evenHeader` the odd
  pages. A kind with no variant shows its default on those pages too. An empty
  variant (`{"blocks": []}`, or a band with no text) blanks that page when the
  kind has a default band, and asks for nothing otherwise. Written as Word's
  `first` / `even` header references (`w:titlePg`, `w:evenAndOddHeaders` in
  settings), parts numbered by slot: `{kind}1.xml` default (unchanged),
  `{kind}2.xml` first, `{kind}3.xml` even, ids `rId{Kind}N`. `read()` returns
  them from any producer: a first (even) variant only while `w:titlePg`
  (`w:evenAndOddHeaders`) is on, never one naming the default's own part, and
  `{"blocks": []}` for a missing or empty variant part while the default has
  content. Footnotes stay refused in every band. Every header, footer, settings
  and document part is byte-identical to PHP's (pinned by the parity suite).
- **Field tokens in header and footer text**: `{section}`, `{sectionpages}`,
  `{author}`, `{title}`, `{filename}`, `{date}`, `{time}`, and
  `{date:FORMAT}` / `{time:FORMAT}` with a Word date-time picture (one or more
  characters, none of `{`, `}`, `"` or a backslash), beside `{page}` /
  `{pages}`. Each is a live complex field, written by the same code `{page}`
  already used, instruction between single spaces (` DATE \@ "FORMAT" ` for a
  picture). The cached result is deterministic: `1` for page and section
  numbers; `meta.author` and the `title` for AUTHOR and TITLE (empty without
  one); no result run for DATE, TIME and FILENAME (no clock: the same document
  still writes the same bytes). Without a picture, DATE and TIME show in the
  reader's locale. Anything else in braces stays literal, and every token stays
  literal in the body. `read()` returns any producer's field as its token,
  matching the first word in any case, keeping a `\@` picture the grammar can
  hold and dropping every other switch. `Schema.FIELD_TOKEN`,
  `Schema.FIELD_INSTRUCTIONS`, `Schema.splitFieldTokens()` and
  `Schema.fieldTokenOf()` are the grammar.
- **`meta.author`**: `docProps/core.xml` `dc:creator`, written after `dc:title`
  and read back (an empty `dc:creator` reads as no `meta`). The title stays the
  top-level `title`; there is no `meta.title` (validation refuses it and
  `validateAndRepair` moves it to `title`, or drops it when there is one), so
  `dc:title` has one slot.
- `Schema.PAGE_BANDS` / `Schema.PAGE_BAND_SLOTS`; the JSON schema's `meta`,
  the four variants and a `blankBand` definition; types `PageBand`, `DocMeta`,
  `FieldToken`, and `Page.header` ... `Page.evenFooter`.
- Agent guide: the first-page band, the letterhead option and the field tokens
  (the guide's bytes and hash changed; the PHP and Python packages carry the
  same bytes). New sha256:
  `6d70d774cf630c0c40813d70ddd00191cc8a0b57a2cd3c603cae54f64a9501fa`.

### Changed

- **`read()` of a document with `w:titlePg` on, or `w:evenAndOddHeaders` on,
  now returns `firstHeader` / `firstFooter` / `evenHeader` / `evenFooter`.**
  Before, only the default band was read and the first/even ones were lost.
  Consumer action: none to keep working; a document read and written back now
  keeps its different first page. If you compared `page` keys exactly, expect
  the new ones.
- **`read()` of a SECTIONPAGES field now returns `{sectionpages}`** (it was
  `{pages}`). Consumer action: none; in a one-section document both show the
  same number, and writing `{sectionpages}` back writes SECTIONPAGES again.
- **A header or footer whose text already contained one of the new tokens
  literally** (say `{date}`) now writes a live field there. Consumer action:
  only if you meant the braces literally; write `{Date}` or another non-token
  spelling.
- `docProps/core.xml` declaring a DOCTYPE is no longer read for the title (or
  the author), as in PHP.

### Fixed

- A field that is not a token, in a one-line header or footer from another
  producer (a STYLEREF, say), lost its cached result on read; it now keeps it
  as text. A `w:fldSimple` that is not a token keeps its text too.

### Notes

- **Not here, because of what this engine does not have yet:**
  - **`templateFirstPageOnly`** (PHP 0.16.0, the house letterhead on the first
    page only). It moves the header and footer that `templateParts` carries
    from a template's section (PHP 0.13.0, last-word#5); this engine takes a
    template's styles and theme only, so the option has nothing to act on. It is
    not in `WriteOptions`, and passing it changes nothing (pinned by a test).
  - **A `{blocks}` header or footer with content** (PHP 0.14.0, last-word#6) is
    still not written: as a first or even variant it is ignored (the default
    shows on that page), as a `{blocks}` `header` always was here; and the
    tokens in a blocks band's run text are not written either. The empty
    `{"blocks": []}` variant IS written (a blank page). Reading, a band is still
    its first paragraph's text, so a first or even part with none there (an
    image-only letterhead) reads as `{"blocks": []}` while the default has
    content, where PHP returns its blocks.
  - The agent guide is shared bytes and so mentions both; an agent following it
    on this engine gets the one-line forms only.
  - Validation paths keep this package's JSON-Pointer form (`/meta/title`,
    `/page/firstFooter/left`).
- `@particle-academy/fancy-conformance` (devDependency) moves to `^0.41.0`:
  `last-word/docx-constructs` rows 0054-0079 pin the letterhead option, the band
  variants, the field tokens and `meta`, with the new functions `bandParts`,
  `bandFields` and `coreProps` and a row input of `{doc, template?, options?}`
  (the template is a document model, written and passed as the template).
  15 of the 26 pass here; the other 11 are skipped by this package's harness,
  each with its reason printed in the run: 0054-0060 need `templateParts`, and
  0061, 0064, 0065 and 0069 write a `{blocks}` band. A skipped row that starts
  passing fails the build until its skip is removed. Dev only; nothing at
  runtime depends on it.
- LibreOffice 26 renders a three-page document written by this package with its
  first-page header on page one only (no footer there, from an empty
  `firstFooter`), the even header on page two and the default on page three, and
  computes PAGE, NUMPAGES, AUTHOR, TITLE and a pictured DATE on load. It shows
  SECTIONPAGES as its cached `1` rather than computing it.

## [0.12.0] - 2026-11-01

**What you must do: nothing, unless you want footnotes or the header/footer
distances.** No existing document writes or reads differently: a document
without a footnote is byte-identical, and `read()` output is unchanged for every
file that has no footnotes and default header/footer distances. Mirrors PHP
`particle-academy/last-word` 0.15.0.

### Added

- **Footnotes** (MOIC, 2026-11-01). A footnote is a RUN placed where the mark
  goes: `{"text": "", "footnote": {"blocks": [...]}}`, its blocks the body's
  vocabulary less `chart` and `pageBreak`. Word numbers the marks in document
  order; the model carries no number. Allowed in headings, paragraphs, list
  items, table cells and quotes; refused in a header or footer and inside
  another footnote (Word has neither). `validateAndRepair` sets a missing `text`
  to `""`, moves text typed on the mark into a run before it, wraps a string or a
  bare list of blocks, and drops a misplaced footnote with a note.
  - `toBytes` writes `word/footnotes.xml` with the separator (`w:id` -1) and
    continuation separator (0) Word requires, `w:footnotePr` naming them in
    `settings.xml`, a `FootnoteReference`-styled `w:footnoteReference` run, the
    relationship and content type, the `FootnoteText` / `FootnoteReference`
    styles (only when a footnote exists; a template's own definitions win), and
    `word/_rels/footnotes.xml.rels` for a footnote's links and images.
    LibreOffice renders the notes at the foot of the page, under the separator
    rule and above the footer, and its own `.docx` export of the file reads back
    with every footnote. The footnotes part, its relationships and settings are
    byte-identical to PHP's (pinned by the parity suite).
  - `read()` returns the same run for a `.docx` footnote from any producer: any
    ids, Word's separators and continuation notices skipped, the space after a
    footnote's own number stripped, a run holding text and a reference split in
    two. A reference whose footnote is missing reads as nothing. A footnotes part
    declaring a DOCTYPE is not read. `mergeRuns()` never merges or drops a
    footnote run.
  - Agent guide: one idiom and example (the guide's bytes and hash changed; the
    PHP and Python packages carry the same bytes).
- **Header and footer distances: `page.margins.header` / `page.margins.footer`**,
  points from the page edge to the header and to the footer (`w:pgMar` `header` /
  `footer`; 36 by default). Written and read symmetrically.
- **`Schema.pageMargins(doc)`**: all four margins and both distances, defaults
  filled in, for a paged viewer. `Schema.PAGE_MARGIN_DEFAULT` (72) and
  `Schema.PAGE_BAND_DISTANCE_DEFAULT` (36) name the defaults, and
  `Schema.FOOTNOTE_REFUSED_TYPES` the block types a footnote refuses. Types:
  `Footnote`, `PageMargins`, `ResolvedPageMargins`; `Run.footnote`.

### Notes

- **The margin rule is unchanged, and now documented.** `read()` omits any margin
  equal to its default (72pt a side, now also 36pt a distance) and omits `page`
  when nothing differs, so "no bottom margin" in a read document means one inch,
  not none. Emitting every side would have changed the read output of every
  existing document and fixture. **Consumer action for a viewer:** resolve with
  `Schema.pageMargins()` (or the same defaults in your language) instead of
  treating an absent margin as zero.
- **Not covered yet:** the `.doc`, `.odt` and `.rtf` readers still skip footnotes
  (as they always have; this release did not change them). Markdown export does
  not carry footnotes: the mark vanishes there and the footnote's text is not
  exported. (This package has no HTML, plain-text or portable export; in PHP
  those do not carry footnotes either.)
- **Where this engine differs from PHP 0.15.0, because of what it does not have
  yet, not because of footnotes:** it writes no `{blocks}` header or footer (PHP
  0.14.0, last-word#6), so a footnote can never reach one here; validation still
  refuses a footnote in a `page.header.blocks` / `page.footer.blocks` and
  `validateAndRepair` drops it, so an agent is told the same thing by both
  engines. And a template's section (`templateParts`, PHP 0.13.0) is not taken,
  so there is no template distance for the document's to override. Validation
  paths keep this package's JSON-Pointer form (`/blocks/0/runs/0/footnote`).
- `@particle-academy/fancy-conformance` (devDependency) moves to `^0.40.0`:
  `last-word/docx-constructs` rows 0045-0053 pin the footnote and page-geometry
  contract for all three engines, with the new `footnotesPart` extractor. Dev
  only; nothing at runtime depends on it.

## [0.11.0] - 2026-10-27

### Added

- **Word `.dotx` and LibreOffice `.ott` templates.** `Agent.read()` reads them
  exactly as their `.docx` / `.odt` twins and adds `sourceKind: "template"`.
  Document reads are unchanged and carry no `sourceKind` key, so nothing to do
  unless you want to tell the two apart. Mirrors PHP `particle-academy/last-word`.
- **`Agent.convertTemplate(bytes, to?)` / `convertTemplate()`** and a
  `last-word` bin: `npx last-word template:convert <in> <out> [--to=document|template] [--force]`
  turn a template into its document twin and back. Only the declaration is
  rewritten: the main part's content-type value (OOXML), or the `mimetype`
  entry plus the manifest root's `media-type` (ODF 1.3 Part 2 section 3.3
  requires them to agree). Every other zip member is copied as its original
  record, byte for byte; `mimetype` stays first and stored. The bin is a
  separate, Node-only build entry (`dist/cli.js`); the package index stays
  isomorphic.
- `OfficeTemplateRefused` with a constant `reason`, and the `SourceKind` type.

### Changed

- **A macro-enabled template (`.dotm`) is now refused** with
  `OfficeTemplateRefused` reason `macro_enabled`; it used to be read as a
  document with its VBA silently ignored. So is a `.dotx` carrying a
  `vbaProject.bin`. A `.docm` still reads as before but is never converted.
  Re-save a `.dotm` as `.dotx` if you need it read.
- A package that declares a template of another family (a PowerPoint or Excel
  template type, an ODF `presentation-template`, a `text-template` whose body is
  a spreadsheet) is refused with `family_mismatch` instead of being read as
  whatever its parts looked like.
- The ESM build now splits shared code into a chunk (`dist/chunk-*.js`) because
  it has two entries. Import paths are unchanged.

## [0.10.1] - 2026-10-26

### Security

- Harden public `unzipSync()`, DOCX/ODT readers and template reads against ZIP
  bombs, duplicate/unsafe names, CRC/size mismatch, inconsistent metadata and
  truncation. Inflation and Node template file input are bounded before allocation.
  `ArchiveRefused.reason` is typed; templates retain it as `TemplateException.reason`.
  `unzipSync()` accepts `ArchiveLimits`; valid fixtures and projections are
  unchanged. No runtime dependency added.

- Update Vitest to 4.1.12 for security fixes; dev toolchain only, nothing to do.

- **Dev toolchain only — no runtime change, nothing for consumers to do.** The
  lockfile now resolves `vite` 8.3.5 (was 8.2.2), whose bundled
  `launch-editor` fixes GHSA-587h-7wmj-h5pm (HIGH, `fs.existsSync()` check
  bypassable on Windows), GHSA-698v-4r3x-3vcv (`escapeCmdArgs()` omits `%`
  escaping on Windows) and GHSA-r825-w8mm-5pxw (editor CLI flag
  misinterpretation). `vite` is only reached through the test/dev toolchain
  and the published package does not ship our lockfile, so the exposure was
  this repo's own dev server.

## [0.10.0] - 2026-10-20

A minor release. Pre-1.0, breaking changes may land in a minor; this one has none.

### Added

- **A native `chart` block** (last-word#4), the same as `particle-academy/last-word`
  0.10.0. Bar (vertical columns or horizontal bars), line and pie charts are
  written as real Word charts: a `word/charts/chartN.xml` part inline in the
  text, with an embedded workbook so Word's "Edit Data" opens the numbers. The
  block takes an Apache ECharts `option`, the same field dark-slide's chart
  element takes, so one prompt teaches charts in both formats:
  `{ type: "chart", option: { xAxis: { type: "category", data: ["Q1", "Q2"] }, yAxis: { type: "value", name: "USD (k)" }, series: [{ type: "bar", name: "Revenue", data: [120, 135] }] }, numberFormat: "#,##0" }`.
  A category `yAxis` gives horizontal bars; a series `stack` stacks bars;
  `label.show` draws data labels; `itemStyle.color` and `option.color` colour a
  series, a slice or one highlighted point; `title`, `legend` and each axis
  `name` are honoured. Sized by `widthPx` / `heightPx`, or `width` as a
  percentage of the text column, or the full column at 0.6 of its width as
  tall. `numberFormat` is an Excel number format; `alt` the alternative text.
  New exported type `ChartBlock`, and the helpers `chartSpec`, `chartNum`,
  `chartLabel`, `chartParseNum`, `SUPPORTED_CHART_TYPES`, `DEFAULT_PALETTE`,
  `DEFAULT_ASPECT`.
- **The reader reads charts back** into the same block, in a canonical form
  (colours on each series, omitted when they are the default palette's;
  stacking as `stack: "total"`; the legend as `{ show: true }`; a percentage
  width as the `widthPx` / `heightPx` it resolved to). A chart it cannot map is
  skipped, never a failed read.
- **Agents learn it from the package**: `Agent.validate()`,
  `Agent.validateAndRepair()` (the `{chartType, categories, series: [{name, values, color}]}`
  shorthand and `series[].values`), `Agent.jsonSchema()`, `Agent.guide()` (the
  guide's hash moves in all three engines) and `Agent.toMarkdown()` (a chart
  renders as its data table) all know the block.
- Not supported, and a validation error: scatter, area and combination charts.

**Nothing to do**: a document without a chart writes the same bytes as before.
The parity suite holds every chart part, the embedded workbook's included,
byte-identical to the PHP writer.

## [0.9.0] - 2026-10-19

### Added

- **An agent guide ships with the package: `agent-guide/professional-documents.md`, and
  `Agent.guide()` returns it.** What makes a document professional and the schema
  idioms that produce it, distilled from rubric-judged agent builds in the
  Fancy document lab: every rule in it moved a measured score. Give it to a
  model together with the JSON schema; it works with any agent loop. The same
  bytes ship in the PHP, Node and Python packages (a test in each pins the
  hash). **Nothing to do**; nothing else changed.

## [0.8.0] - 2026-10-19

### Added

- **Running headers and footers with live page numbers**: `page.header` and
  `page.footer`, each one line with `left`, `center` and `right` slots (plus
  optional `size` and `color`); `{page}` and `{pages}` become live PAGE and
  NUMPAGES fields. The reader reads them back. **Nothing to do**: a document
  without them is byte-identical. Same as `particle-academy/last-word` 0.8.0 and
  `fancy-last-word` 0.4.0; parity holds (new `bands` parity fixture).

### Fixed

- `package-lock.json` carried `0.6.1` through the 0.7.0 release; it now matches.

## [0.7.0] - 2026-10-07

### Added

- **`{ template }` on `Agent.toBytes()` and `Agent.write()` — render onto a house
  template** (last-word#3), matching PHP `last-word` 0.7.0. The template's
  `word/styles.xml` and `word/theme/theme1.xml` are carried verbatim; definitions
  the template lacks — `CodeBlock`, `InlineCode`, heading levels it omits — come
  from the built-in set.

  **Binds by style name, with nothing to configure**, because the document model
  already uses Word's own style ids. Two Word behaviours shape it, and both fail
  SILENTLY rather than loudly: a `w:pStyle` naming an undefined style renders
  unstyled, so a template without `CodeBlock` would quietly produce unformatted
  code; and a duplicate definition is legal with the LATER one winning, so
  appending a style the template already has would override the house style for
  exactly the subset the two happen to share. Hence: append only what is missing.

  The theme travels with the styles deliberately — a style naming a theme colour
  or font resolves against whatever theme is in the package, so styles without
  their theme give the template's structure in the default's colours.

  **Not taken from the template:** `w:sectPr` (page size, margins, headers,
  footers), `word/numbering.xml` (this document's lists reference numbering ids
  defined here) and `word/settings.xml`. A cover page is not in this release.

  `toBytes()` takes the template as **bytes** — it is universal and a browser has
  no filesystem — and says so rather than failing with a confusing zip error.
  `write()` accepts bytes or a path.

  **Nothing changes for callers who do not pass it.** The built-in look is
  byte-for-byte what it was, which the determinism suite and the PHP parity suite
  both check.

- **`TemplateException`** (exported), thrown when a template is not a readable
  package or has no `word/styles.xml`. It REFUSES rather than falling back,
  because a document that silently comes out in the built-in style is the exact
  failure the option exists to end — and it gives a host a way to validate a
  customer-supplied template at upload rather than at render.

- **`WriteOptions`** is exported as a type.

### Fixed

- **The parity suite's `word/styles.xml` divergence entry named one cause and
  there are three.** Found while reading both engines' style sets side by side for
  the template work. It said only `w:eastAsia`; it now also records that
  **`Heading1` is 40 half-points here and 36 in PHP** (user-visible — the same
  document gets a different top heading size from each engine) and that **PHP
  defines `Title` and `ListParagraph` where this port defines neither**.

  Behaviour is unchanged: reconciling either changes rendered output for every
  existing consumer, which is the pair owner's call. But a divergence ledger is the
  one place anyone looks before trusting the pair, and an entry that understates is
  the thing that file exists to prevent. It matters for templates specifically: a
  house template's "Title" and "List Paragraph" styles bind in PHP and have
  nothing to bind to here.

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
