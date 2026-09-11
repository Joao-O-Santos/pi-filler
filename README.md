# pi-filler

<p align="center"><img src="logo.png" alt="pi-filler logo" width="220"></p>

Deterministic DOCX and PDF tooling for Pi.

`pi-filler` fills document-handling gaps between ordinary Markdown
workflows and final office-format requirements. It delegates common
conversion and PDF extraction to mature command-line tools while keeping
Word-specific mutations narrow and inspectable.

## Status

The v0.0.0 implementation is a release candidate. It has not been
tagged or released.

## Tool

The extension exposes one `filler` tool:

```text
format: docx | pdf
action: read | search | write | patch
view: text | formatting | image
```

Unsupported combinations and operation-irrelevant parameters fail
explicitly. DOCX patching also accepts `dry_run: true` to validate and
report the requested result without writing the output file.

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
- Inspect page setup, margins, sections, line numbering, page
  numbering, comments, inserted/deleted tracked-change markers, and
  common core metadata.
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

## External tools

Required for the corresponding operations:

- `pandoc` for DOCX text conversion and generation.
- `libreoffice` for DOCX image rendering.
- `pdftotext` for PDF text extraction.
- `pdftocairo` for PDF and DOCX page rendering.
- `pdfgrep` for page-aware PDF search.

## Development

```sh
npm install
make verify
make site
```

Dependencies and CI actions follow current upstream releases rather than
being pinned. GitLab is canonical and the only release authority.
GitHub may mirror the repository, verify it, and publish GitHub Pages.
