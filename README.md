# pi-filler

![pi-filler logo](logo.png)

[![pipeline
status](https://gitlab.com/Joao-O-Santos/pi-filler/badges/main/pipeline.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/commits/main)
[![npm
version](https://img.shields.io/npm/v/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![npm
downloads](https://img.shields.io/npm/dt/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![license](https://img.shields.io/npm/l/pi-filler.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/blob/main/LICENSE)

**Deterministic DOCX, PDF, PPTX, and XLSX tooling for Pi.**

`pi-filler` covers document tasks that do not fit an ordinary Markdown
workflow. It delegates conversion and PDF extraction to mature
command-line tools while keeping DOCX, PPTX, and XLSX mutations narrow,
transactional, and inspectable.

The project is experimental and pre-1.0. The README documents the
current main branch; the npm badge shows the currently published
version, which may not yet contain every feature described here.

## Install

Install the published package:

``` sh
pi install npm:pi-filler
```

Or install the current main branch for development and unreleased
features:

``` sh
pi install git:https://gitlab.com/Joao-O-Santos/pi-filler.git
```

External programs are required only for the operations listed under
[External tools](#external-tools). XLSX operations run locally through
Node.js and do not require an office suite.

## Tool and operation matrix

The extension exposes one `filler` tool:

``` text
format: docx | pdf | pptx | xlsx
action: read | search | write | patch
view: text | formatting | image | structure | review
```

`format` names the document or target format, not every input: DOCX and
PPTX writes take Markdown at `path`, and XLSX writes take a workbook
template at `path` plus CSV at `source_csv`. Choose a supported
combination and supply only its relevant fields:

| Format | Action | View | Required additional fields |
|---------------|---------------|---------------|---------------------------|
| DOCX | `read` | `text` | optional `track_changes` |
| DOCX | `read` | `formatting` | none |
| DOCX | `read` | `review` | optional `review_offset`, `review_limit` |
| DOCX | `read` | `image` | `output`; optional `first_page`, `last_page` |
| DOCX | `search` | `text` or omitted | `query`; optional `ignore_case`, `track_changes` |
| DOCX | `write` | omitted | Markdown `path`, `output`; optional `reference_docx` |
| DOCX | `patch` | omitted | `output` and exactly one of `patches` or `tracked_replace`; optional `dry_run` only with `patches` |
| PPTX | `read` | `text` or `formatting` | none |
| PPTX | `read` | `image` | `output`; optional `first_page`, `last_page` |
| PPTX | `search` | `text` or omitted | `query`; optional `ignore_case` |
| PPTX | `write` | omitted | Markdown `path`, `output`; optional `reference_pptx` |
| PPTX | `patch` | omitted | `output`, `pptx_patches`; optional `dry_run` |
| PDF | `read` | `text` | optional `first_page`, `last_page` |
| PDF | `read` | `image` | `output`; optional `first_page`, `last_page` |
| PDF | `search` | `text` or omitted | `query`; optional `first_page`, `last_page`, `literal`, `ignore_case` |
| XLSX | `read` | `structure` | none |
| XLSX | `read` | `image` | `output`; optional `first_page`, `last_page` select resulting workbook print pages |
| XLSX | `write` | omitted | template `path`, `source_csv`, `sheet`, A1 `start_cell`, `output`; optional `column_map` |
| XLSX | `patch` | omitted | `output` and either `sheet` with A1 `range`/`xlsx_patches`, `sheet` with `xlsx_formula` / `xlsx_formulas`, or `xlsx_formula_sheets` without `sheet`; optional `dry_run` |

PDF write and patch, and XLSX search, are unsupported. Parameters
irrelevant to the chosen operation fail explicitly instead of being
ignored.

Paths resolve relative to Pi's current working directory; absolute paths
and Pi-style paths prefixed with `@` are also accepted. A write or patch
output must differ from its input. For image reads, `output` is an
existing directory or PNG prefix; the operation creates page images
rather than one final PNG file.

`first_page` and `last_page` select 1-based inclusive pages. Both must
be positive, and `first_page` cannot exceed `last_page`. Use
`start_cell` like `A2` and `range` like `A2:D10`; a range may also be
one cell.

PDF search interprets `query` as a regular expression by default; set
`literal: true` for literal text. It is case-sensitive unless
`ignore_case: true` is set. XLSX writes skip the first CSV record by
default and preserve fields as text by default. `dry_run: true` still
requires `output`, but it does not create or replace it.

Example calls:

``` ts
filler({
  format: "pdf",
  action: "search",
  path: "paper.pdf",
  query: "registered report",
  ignore_case: true
})

filler({
  format: "xlsx",
  action: "write",
  path: "template.xlsx",
  source_csv: "results.csv",
  sheet: "Results",
  start_cell: "A2",
  output: "completed.xlsx",
  dry_run: true
})
```

Results may be truncated. For a truncated PDF text read, use a page
range; for PDF, DOCX, or PPTX search, narrow the query. PDF search can
also use a page range. Dry runs still read and validate all local
inputs.

## PDF operations

- Text reads use layout-preserving extraction and optional page ranges.
- Search returns page-aware matches and supports 1-based inclusive page
  ranges, case-insensitive matching, and literal rather than
  regular-expression queries.
- Image reads render selected pages to numbered PNG files.
- A failed render preserves previous files for the requested output
  prefix.

## DOCX operations

- Text reads and search use Pandoc Markdown, not GFM. `track_changes`
  defaults to `accept`; choose `reject` for pre-edit text or `all` for
  insertion/deletion spans and comments. This selects an extraction
  view; it does not perform native Word acceptance or rejection.
- Image reads convert DOCX through LibreOffice and then render the
  resulting PDF pages.
- Writes convert Pandoc Markdown into a validated DOCX, including
  supported native review spans and anchored comments. Smart typography
  is disabled, preserving literal quotes and dashes. An optional
  reference DOCX supplies styles and layout defaults; conversion is not
  a lossless round trip of existing Word layout, comment threads, or
  every revision type.
- Formatting reads inspect page setup, margins, sections, line and page
  numbering, comments, inserted/deleted tracked-change markers, and
  common core metadata.
- Review reads return comment IDs, bounded text/author/date fields,
  main-document anchor excerpts, and supported legacy reply/resolution
  metadata. They never edit comments or resolve threads.
- Patches change first-section page setup, margins, line numbering, page
  numbering, and common core metadata. Alternatively, `tracked_replace`
  makes a narrow tracked text replacement using an optional separate
  executable.

### Tracked replacement

Use `format: "docx"`, `action: "patch"`, `path`, `output`, and exactly
one `tracked_replace` object instead of `patches`. Omit `view` and
`dry_run` (even `false` is rejected). The object accepts only:

- `old`: nonempty string containing the literal text to match;
- `new`: string containing replacement text; an empty string deletes;
- `author`: nonempty string for the revision author;
- `expected_sha256`: optional lowercase, 64-hex-digit source hash.

Install `docxquisite` separately and make the executable available on
`PATH`; pi-filler never installs or downloads it. For example, from an
accessible toolkit checkout, run `uv tool install /path/to/docxquisite`
and ensure the uv tool executable directory is on `PATH`.
[Docxquisite](https://gitlab.com/Joao-O-Santos/docxquisite) is MIT
licensed. The output must differ from the input and must not already
exist. Publication requires hard-link support in the output directory.
Pi-filler probes protocol v1 capability, computes the source SHA-256,
checks any supplied expected hash, and passes its computed hash to the
toolkit for freshness checks. The final check is not a lock against
concurrent changes. On success, the result contains `output`,
`source_sha256`, and `output_sha256`; pi-filler independently hashes the
output and checks it against the toolkit's result.

Replacement requires one unique literal match within a single text run
in a simple main-document paragraph or table. Existing basic insertions,
deletions, and unthreaded comments may be retained outside the target
paragraph under strict structural checks, with no comment range
enclosing the target. Threaded comments, complex revisions, fields,
other stories, cross-run matches, signed packages, and other unsupported
structures are rejected. This does not accept or reject revisions or
edit comments. Structural validation does not establish Microsoft Word
behavior or visual fidelity; inspect a copy in the target application.

The request, including text and author, is sent locally over stdin, not
command-line arguments. It remains sensitive in process memory and
pipes; protect these and the output file. Normal results omit the
requested text, author, and toolkit diagnostics. A transport,
subprocess, or protocol failure during replacement can leave the output
in an unknown state. Inspect any output before retrying; pi-filler does
not automatically retry or clean it up. Cancellation and timeout
terminate the direct child only, not necessarily its descendants.

### Review inspection and formatting patches

Review pagination uses zero-based `review_offset` (default 0) and
`review_limit` (default 20, maximum 50). A byte budget may return fewer
records; follow `nextOffset` until it is null. Comment text is capped at
1,000 Unicode code points and anchor excerpts at 200, with truncation
flags; author/date fields are bounded to 128/64 UTF-16 code units.
Ranges include insertion and deletion text, not an accepted-view
projection. Point anchors show paragraph context. A missing anchor means
not found in the main document, not necessarily a damaged comment.
Legacy `commentsExtended` metadata supports resolution and reply-parent
reporting where matched; additional modern thread features may be
unsupported and are reported as warnings. This is not native Word
acceptance/rejection or a complete review-history export.

DOCX page dimensions and margins use twips (twentieths of a point); 1
inch is 1,440 twips. Page setup, margins, line numbering, and page
numbering patches apply to the first section only. Core-metadata patches
apply to the document.

`clearCoreMetadata` clears common `docProps/core.xml` fields. It does
**not** remove comment authors, revision authors, or every possible
identity-bearing property and must not be described as anonymization.

Untouched ZIP package parts remain content-identical. Edited XML parts
are parsed and serialized again, so their lexical formatting, prefixes,
or element ordering may change.

## XLSX operations and privacy

Existing spreadsheet cell values, CSV fields, formula contents, and
comments are processed locally and are not returned to the model. A
formula supplied in `xlsx_formula`, `xlsx_formulas`, or
`xlsx_formula_sheets` is tool input and visible to the agent. When a
deterministic XLSX write or patch does not require inspection, do not
read CSV or cell contents into model context merely to perform it.
Structural results may contain sheet names, ranges, dimensions, counts,
merged-range addresses, hidden row and column ranges, changed package
parts, and write status---not cell or CSV contents. Layout-range arrays
are capped and report whether they were truncated; this metadata can
reveal limited layout and occupancy information.

- `read` with `view: "structure"` reports workbook and worksheet shape.
- `read` with `view: "image"` renders workbook print pages through
  LibreOffice to numbered PNG files; it does not select a sheet.
  `first_page` and `last_page` select resulting print pages.
- `write` fills an existing workbook template from CSV.
- `patch` changes formatting or layout while preserving cell contents,
  or sets explicit ordinary formulas atomically.

CSV writes skip the first record by default (`has_header: true`) and
preserve fields as text by default (`value_mode: "text"`). Automatic
mode recognizes only empty values, numbers, and booleans; it does not
infer dates or formulas. `column_map` optionally maps either all 1-based
CSV ordinals or all exact CSV headers to absolute XLSX column letters,
such as `{ "1": "A", "2": "C" }`. Header mappings require unique CSV
headers; destinations must be unique. The entire write fails if its
mapped destination cells intersect an existing formula or merged range.

XLSX patches support bold, italic, font size, six-digit RGB fills,
horizontal and vertical alignment, wrapping, Excel number-format codes,
uniform borders, column width, and row height. Output is validated and
atomically replaced only after successful processing.

### Adding or editing formulas

Use `path` for the input workbook, `sheet` for its worksheet name,
`output` for a distinct output path, and exactly one of `xlsx_formula`
(single cell) or `xlsx_formulas` (atomic batch), instead of `range` and
`xlsx_patches`:

``` ts
filler({
  format: "xlsx",
  action: "patch",
  path: "template.xlsx",
  sheet: "Results",
  output: "with-formula.xlsx",
  xlsx_formulas: [
    { cell: "B2", formula: "=SUM(B3:B10)" },
    { cell: "C2", formula: "=B2*2" }
  ],
  dry_run: true
})
```

A batch accepts 1--1,000 explicit targets on one sheet. Duplicate
case-normalized cells (such as `b2` and `B2`), empty or oversized
batches, and conflicting patch fields are rejected. All targets are
validated before publishing one output; a failure leaves existing output
unchanged. Absent cells are created; existing values or formulas are
replaced, preserving cell styles. There is no fill-down or
relative-reference translation: supply each formula explicitly. Batch
results include normalized `cells` in request order and `formulaCount`;
single-cell results retain `cell`. Neither returns formulas or values.

For an atomic edit across worksheets, use `xlsx_formula_sheets` instead
of `sheet`, `xlsx_formula`, `xlsx_formulas`, `range`, and
`xlsx_patches`:

``` ts
filler({
  format: "xlsx",
  action: "patch",
  path: "template.xlsx",
  output: "with-formulas.xlsx",
  xlsx_formula_sheets: [
    { sheet: "Results", formulas: [{ cell: "B2", formula: "=SUM(B3:B10)" }] },
    { sheet: "Summary", formulas: [{ cell: "A1", formula: "=Results!B2" }] }
  ]
})
```

Each group requires exactly `sheet` and nonempty `formulas`. There must
be 1--1,000 targets total, with unique normalized cells per sheet and no
repeated worksheet targets. Formula worksheet names resolve
case-insensitively; aliases resolving to the same worksheet part are
rejected, as are ambiguous workbook names or part mappings. Every sheet
and target is checked before mutation, using one source load and one
atomic output publication, not sequential file edits. Dry runs validate
all sheets without writing. Results contain `sheets` in request order,
each with the workbook's `sheet` name, normalized `cells`, and
`formulaCount`, plus a total `formulaCount` and shared output metadata.
No formula or value contents are returned. Unrelated package parts
remain byte-identical at the part content level; edited worksheet XML is
reserialized. Existing single-sheet request and result shapes remain
supported.

`cell` must be one explicit, unmerged A1 cell. `formula` must start with
`=`, have a nonblank body, contain XML-safe characters without carriage
returns (CR), and be at most 8,192 UTF-16 code units long. Only `cell`
and `formula` are accepted in the object. The target's old value or
formula and cached value are removed; their contents are not returned.
The supplied formula is not evaluated or semantically validated. Other
formula caches remain unchanged and may be stale. Workbook flags request
automatic full recalculation on open once per operation; they do not
perform or guarantee recalculation. In one synthetic headless
LibreOffice 26.8.2.2 reopen/export, the edited cell recalculated but its
dependent cell retained a stale cached value. Verify recalculation in
the target application before relying on results.

Formula edits discard a conventional `xl/calcChain.xml` calculation
order cache and its workbook relationship and content-type override.
`changedParts` includes the removed part and edited metadata. This also
happens during validation for a dry run, without writing any files.
Nonstandard, orphaned, or ambiguously linked chains are rejected. Chain
removal does not calculate formulas or clear dependent cell caches.

Formula editing rejects signed packages, shared, array, and data-table
formulas, unsupported extension or formula metadata, ambiguous sheet,
row, cell or relationship structures, and other unsupported workbook
structures. It validates a temporary sibling package before atomic
replacement of the distinct output; the input is not intentionally
modified. `dry_run: true` validates without creating or replacing
output.

Formulas can invoke external references or other potentially unsafe
behavior when a spreadsheet application later opens the workbook. Treat
formula input and output as untrusted where appropriate; review them
before opening in an application that evaluates formulas.

XLSX support does not provide cell-content reads or search, formula
evaluation or automatic formula generation, spreadsheet scripting, or
manipulation of charts, pivots, macros, comments, named ranges, tables,
data validation, conditional formatting, or external data.

## PPTX operations

- Text reads and search convert PPTX through Pandoc to GFM Markdown.
- Writes convert GFM Markdown into a validated PPTX; an optional
  reference PPTX supplies Pandoc presentation styles.
- Formatting reads report slide dimensions, slide count, and common core
  metadata.
- Patches change slide dimensions, common core metadata, and targeted
  text in slide text paragraphs. Text replacements may target one
  1-based slide or all slides. They can span formatted text runs, with
  replacement text inheriting the first run's formatting.
- PPTX-to-Markdown extraction is semantic and lossy: it is not a
  reversible representation of positions, animations, themes, or all
  presentation objects.

Slide dimensions use EMUs (English Metric Units); 914,400 EMUs equal one
inch. PPTX image reads convert slides through LibreOffice to PDF and
render numbered PNG files; `first_page` and `last_page` select slides.

## Output and mutation behavior

DOCX, PPTX, and XLSX writes and ordinary patches use temporary sibling
files, validate the result, and rename it into place only after success.
Tracked replacement instead requires a nonexistent output and uses the
separate toolkit's publication mechanism described above. Source files
are not intentionally modified. PDF and office image rendering similarly
protects prior output until the new render succeeds.

Where supported, `dry_run` validates the requested operation and inputs
without creating or replacing output; its required `output` path is
still validated. It is not supported for tracked replacement.

## External tools

Required for the corresponding operations:

- `pandoc` for DOCX/PPTX text conversion, search, and generation;
- `libreoffice` for DOCX, PPTX, and XLSX image rendering (each
  conversion uses a private temporary profile, removed on completion);
- `pdftotext` for PDF text extraction;
- `pdftocairo` for PDF, DOCX, PPTX, and XLSX page rendering;
- `pdfgrep` for page-aware PDF search; and
- separately installed `docxquisite` on `PATH` for DOCX tracked
  replacement only.

A missing executable produces an explicit error naming the unavailable
program.

## Development and release

``` sh
npm install
make verify
make site
```

`make verify` runs typechecking, lint and formatting checks,
deterministic tests, project-wide line/function coverage checks, and
`npm pack --dry-run`. `make site` regenerates the tracked static site
after documentation changes.

Dependencies and CI actions follow current upstream releases rather than
being pinned. GitLab is canonical; GitHub may mirror the repository,
verify it, and publish GitHub Pages.

## License

[MIT](LICENSE)
