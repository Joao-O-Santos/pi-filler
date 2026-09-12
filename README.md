# pi-filler

![pi-filler logo](logo.png)

[![pipeline
status](https://gitlab.com/Joao-O-Santos/pi-filler/badges/main/pipeline.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/commits/main)
[![npm
version](https://img.shields.io/npm/v/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![npm
downloads](https://img.shields.io/npm/dt/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![license](https://img.shields.io/npm/l/pi-filler.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/blob/main/LICENSE)

**Deterministic DOCX, PDF, and XLSX tooling for Pi.**

`pi-filler` covers document tasks that do not fit an ordinary Markdown
workflow. It delegates conversion and PDF extraction to mature
command-line tools while keeping DOCX and XLSX mutations narrow,
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
format: docx | pdf | xlsx
action: read | search | write | patch
view: text | formatting | image | structure
```

Choose a supported combination and supply only its relevant fields:

| Format | Action | View | Required additional fields |
|--------------|--------------|--------------|------------------------------|
| DOCX | `read` | `text` or `formatting` | none |
| DOCX | `read` | `image` | `output`; optional page range |
| DOCX | `search` | `text` or omitted | `query`; optional `ignore_case` |
| DOCX | `write` | omitted | Markdown `path`, `output`; optional `reference_docx` |
| DOCX | `patch` | omitted | `output`, `patches`; optional `dry_run` |
| PDF | `read` | `text` | optional page range |
| PDF | `read` | `image` | `output`; optional page range |
| PDF | `search` | `text` or omitted | `query`; optional page range and search flags |
| XLSX | `read` | `structure` | none |
| XLSX | `write` | omitted | template `path`, `source_csv`, `sheet`, `start_cell`, `output` |
| XLSX | `patch` | omitted | `sheet`, `range`, `xlsx_patches`, `output` |

PDF write and patch, and XLSX search, are unsupported. Parameters
irrelevant to the chosen operation fail explicitly instead of being
ignored.

Paths resolve relative to Pi's current working directory; absolute paths
and Pi-style paths prefixed with `@` are also accepted. A write or patch
output must differ from its input. Image `output` values are filename
prefixes rather than single image paths.

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

Use `dry_run: true` to validate a DOCX patch or XLSX write/patch without
writing its output. Dry runs still read and validate all local inputs.

## PDF operations

- Text reads use layout-preserving extraction and optional page ranges.
- Search returns page-aware matches and supports optional page ranges,
  case-insensitive matching, and literal rather than regular-expression
  queries.
- Image reads render selected pages to numbered PNG files.
- A failed render preserves previous files for the requested output
  prefix.

## DOCX operations

- Text reads and search convert DOCX through Pandoc.
- Image reads convert DOCX through LibreOffice and then render the
  resulting PDF pages.
- Writes convert a Markdown input into a validated DOCX, optionally
  using a reference DOCX for styles and layout.
- Formatting reads inspect page setup, margins, sections, line and page
  numbering, comments, inserted/deleted tracked-change markers, and
  common core metadata.
- Patches change first-section page setup, margins, line numbering, page
  numbering, and common core metadata.

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

Spreadsheet cell values, CSV fields, formulas, comments, and other
workbook text are processed locally and are not returned to the model.
Structural results contain sheet names, ranges, dimensions, counts,
changed package parts, and write status.

- `read` with `view: "structure"` reports workbook and worksheet shape.
- `write` fills an existing workbook template from CSV.
- `patch` changes formatting or layout while preserving cell contents.

CSV writes skip the first record by default (`has_header: true`) and
preserve fields as text by default (`value_mode: "text"`). Automatic
mode recognizes only empty values, numbers, and booleans; it does not
infer dates or formulas. The entire write fails if its destination
intersects an existing formula.

XLSX patches support bold, italic, font size, six-digit RGB fills,
horizontal and vertical alignment, wrapping, Excel number-format codes,
uniform borders, column width, and row height. Output is validated and
atomically replaced only after successful processing.

XLSX support does not provide cell-content reads or search, formula
calculation or generation, spreadsheet scripting, or manipulation of
charts, pivots, macros, comments, named ranges, tables, data validation,
conditional formatting, or external data.

## Output and mutation behavior

DOCX and XLSX writes and patches use temporary sibling files, validate
the result, and rename it into place only after success. Source files
are not modified. PDF image rendering similarly protects prior output
until the new render succeeds.

`dry_run` validates the requested operation and inputs without writing
output.

## External tools

Required for the corresponding operations:

- `pandoc` for DOCX text conversion, search, and generation;
- `libreoffice` for DOCX image rendering;
- `pdftotext` for PDF text extraction;
- `pdftocairo` for PDF and DOCX page rendering; and
- `pdfgrep` for page-aware PDF search.

A missing executable produces an explicit error naming the unavailable
program.

## Development and release

``` sh
npm install
make verify
make site
```

`make verify` runs typechecking, lint and formatting checks,
deterministic tests, and `npm pack --dry-run`. `make site` regenerates
the tracked static site after documentation changes.

Dependencies and CI actions follow current upstream releases rather than
being pinned. GitLab is canonical and the only release authority. GitHub
may mirror the repository, verify it, and publish GitHub Pages. A
`vX.Y.Z` GitLab tag must match `package.json`; after verification it
publishes to npm with provenance.

## License

[MIT](LICENSE)
