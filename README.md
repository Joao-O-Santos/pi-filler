# pi-filler

Deterministic DOCX and PDF tooling for Pi.

`pi-filler` fills the document-handling gaps between ordinary Markdown
workflows and final office-format submission requirements. It delegates
what mature tools already do well and keeps Word-specific mutations
small, typed, and inspectable.

The core uses Pandoc for DOCX conversion, Poppler tools for PDF text and
rendering, and `pdfgrep` for page-aware PDF search. DOCX formatting
changes operate directly on OOXML rather than relying on a full office
suite.

## Status

v0.0.0 is complete and ready for release. The extension reads and
searches PDFs and DOCX files, renders PDF pages, generates DOCX from
Markdown with reference documents, inspects DOCX formatting, and applies
typed formatting patches. All operations use external command-line tools
or Pandoc with bounded output and transactional writes.

See `ARCHITECTURE.md` for the tool surface and `CHANGELOG.md` for
release notes.

## Supported operations

**PDF:** - Read and extract text with optional page ranges - Search with
literal or regular-expression queries and page-aware results - Render
selected pages to PNG images

**DOCX:** - Read and extract text through Pandoc - Search extracted text
(case-insensitive, line-based) - Generate from Markdown with optional
reference.docx - Inspect formatting state (page size/orientation,
margins, sections, line numbering, page numbering, comments, tracked
changes, core metadata) - Apply typed formatting patches: page setup,
margins, line numbering mode, page numbering start/format, core
metadata, anonymization - Preserve all untouched package parts

**Constraints:** - Output is bounded to 50KB and 2000 lines - Writes and
patches require distinct output paths - Patches include dry-run
support - Mutations are transactional via temporary files

## External tool requirements

Required: - `pandoc` --- DOCX read/write conversions - `pdftotext` ---
PDF text extraction - `pdftocairo` --- PDF page rendering - `pdfgrep`
--- PDF page-aware search

Optional: - `libreoffice` --- future DOCX-to-PDF rendering

## Known limitations

- DOCX patches affect the first section only
- Line numbering modes are limited to `off`, `continuous`, `newPage`,
  `newSection`
- OOXML namespace handling preserves only existing prefixes in
  serialized output; complex documents with alternate namespaces may
  require manual review
- PDF write and patch operations are deferred

## Development

``` sh
npm install
make verify
make site
```

GitLab is canonical and is the only release authority. GitHub may mirror
the repository, verify it, and publish GitHub Pages.
