# pi-filler

<p align="center">
`<img src="logo.png" alt="pi-filler logo" width="220">`{=html}
</p>

[![pipeline
status](https://gitlab.com/Joao-O-Santos/pi-filler/badges/main/pipeline.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/commits/main)
[![npm
version](https://img.shields.io/npm/v/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![npm
downloads](https://img.shields.io/npm/dt/pi-filler.svg)](https://www.npmjs.com/package/pi-filler)
[![license](https://img.shields.io/npm/l/pi-filler.svg)](https://gitlab.com/Joao-O-Santos/pi-filler/-/blob/main/LICENSE)

Deterministic DOCX, PDF, and XLSX tooling for Pi.

`pi-filler` fills document-handling gaps between ordinary Markdown
workflows and final office-format requirements. It delegates common
conversion and PDF extraction to mature command-line tools while keeping
Word- and spreadsheet-specific mutations narrow and inspectable.

## Status

Early experimental release. The tool schema and supported document
operations may change before 1.0.0.

## Tool

The extension exposes one `filler` tool:

``` text
format: docx | pdf | xlsx
action: read | search | write | patch
view: text | formatting | image | structure
```

Unsupported combinations and operation-irrelevant parameters fail
explicitly. DOCX patches and XLSX writes or patches accept
`dry_run: true` to validate and report the requested result without
writing the output file.

## PDF

- Read text with layout preservation and optional page ranges.
- Search with literal or regular-expression queries and page numbers.
- Restrict searches to optional page ranges.
- Render selected pages to PNG files without destroying prior output
  when rendering fails.
- PDF writing and patching are not supported.

## DOCX

- Read and search text through Pandoc.
- Render selected pages to PNG through LibreOffice and the existing PDF
  rendering path.
- Generate DOCX from Markdown with an optional reference document.
- Validate generated DOCX before replacing the requested output.
- Inspect page setup, margins, sections, line numbering, page numbering,
  comments, inserted/deleted tracked-change markers, and common core
  metadata.
- Patch page setup, margins, line numbering, page numbering, and common
  core metadata.
- Use `clearCoreMetadata` to clear common `docProps/core.xml` fields.
  This does **not** remove comment authors, revision authors, or every
  possible identity-bearing property in a DOCX.

Page setup, margins, line numbering, and page numbering patches affect
the first section only. Core-metadata patches are document-level.
Untouched ZIP package parts remain content-identical. XML parts that are
edited are parsed and serialized again, so lexical formatting, prefix
choices, or element ordering inside those edited parts may change.

## XLSX

Spreadsheet contents are processed locally and are not returned to the
model. This includes cell and CSV values, formula expressions, comments,
and other workbook text. Results contain sheet names, ranges, counts,
changed package parts, and write status only.

For spreadsheet work:

- use `read` with `view: "structure"` to inspect workbook shape;
- use `write` with `source_csv`, `sheet`, and `start_cell` to fill an
  existing template;
- use `patch` with `sheet`, `range`, and `xlsx_patches` to change
  formatting or layout.

CSV writes skip the first record by default (`has_header: true`) and
preserve every field as text by default (`value_mode: "text"`).
Automatic mode recognizes only empty values, numbers, and booleans. It
does not infer dates or formulas. A write fails if its destination
intersects an existing formula.

The initial patch surface supports bold, italic, font size, six-digit
RGB fills, horizontal and vertical alignment, wrapping, number formats,
uniform borders, column width, and row height. Cell contents are
preserved. Unsupported package parts remain untouched, and output is
validated and atomically replaced after successful processing.

XLSX support does not provide cell-content reads or search, formula
calculation or generation, spreadsheet scripting, or manipulation of
charts, pivots, macros, comments, named ranges, tables, or data
validation.

## External tools

Required for the corresponding operations:

- `pandoc` for DOCX text conversion and generation.
- `libreoffice` for DOCX image rendering.
- `pdftotext` for PDF text extraction.
- `pdftocairo` for PDF and DOCX page rendering.
- `pdfgrep` for page-aware PDF search.

## Development

``` sh
npm install
make verify
make site
```

Dependencies and CI actions follow current upstream releases rather than
being pinned. GitLab is canonical and the only release authority. GitHub
may mirror the repository, verify it, and publish GitHub Pages.
