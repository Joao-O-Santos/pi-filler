# Changelog

## Unreleased

### Added

- PDF text extraction with layout preservation and page ranges.
- Page-aware PDF search with literal and regular-expression modes.
- Selected-page PDF rendering to PNG.
- DOCX text extraction and search through Pandoc.
- Transactional DOCX generation with optional reference documents.
- DOCX formatting inspection for submission-relevant properties.
- Conservative first-section OOXML patching with dry-run support.
- Relationship and generated-DOCX validation.
- Deterministic regression tests for command limits and document
  operations.
- Structural-only XLSX inspection without spreadsheet-content output.
- Transactional CSV-to-XLSX template filling in text and automatic
  primitive modes.
- Formula-cell protection for XLSX destinations.
- Narrow XLSX style, column-width, and row-height patches with dry-run
  support.
- Privacy regression fixtures covering workbook cells, formulas,
  comments, and CSV fields.

### Changed

- Clarify installation, operation selection, path roles, defaults,
  units, privacy boundaries, and release behavior in user and
  architecture documentation.
- Expand the model-facing schema and prompt guidance so valid
  format/action/view combinations can be selected without trial calls.
- Clarify model-facing page ranges, A1 spreadsheet addresses, image
  output roles, dry-run semantics, defaults, result truncation, and the
  XLSX privacy boundary.
- Search DOCX against the full extraction before truncating results.
- Bound search tool details to match counts and truncation metadata
  instead of duplicating full match arrays.
- Use Node.js `execFile` for bounded command execution, cancellation,
  and finite timeouts.
- Force-terminate timed-out commands so a child ignoring `SIGTERM`
  cannot outlive the configured timeout.
- Honor PDF page ranges during search and use pdfgrep's native
  fixed-string mode for literal queries.
- Reject malformed pdfgrep output instead of assigning synthetic page
  numbers.
- Preserve prior PDF render outputs when a replacement render fails.
- Reject unsupported views, operation-irrelevant parameters, unknown
  parameters, and empty paths instead of silently accepting them.
- Resolve helper input paths consistently before invoking external
  commands.
- Validate XML syntax instead of relying on tolerant parsing.
- Reject empty DOCX patch requests and incompatible line-number
  settings.
- Keep page dimensions consistent when orientation alone is changed.
- Treat omitted page orientation as OOXML's portrait default.
- Count inserted and deleted revision markers in common WordprocessingML
  content parts, including headers, footers, footnotes, and endnotes.
- Preserve attributes on core-metadata elements when their text changes.
- Standardize line-number restart modes on OOXML values.
- Rename the misleading anonymization option to `clearCoreMetadata`.
- Make Word-attribute handling tolerant of alternate namespace prefixes.
- Remove stale PDF render outputs and sort generated pages numerically.
- Follow current releases for runtime dependencies.
- Check Markdown prose width directly instead of requiring a Pandoc
  round-trip to be byte-identical.
- Resolve worksheets through workbook relationships and preserve
  untouched XLSX package parts.
- Keep spreadsheet contents tool-local and return only structural XLSX
  metadata.

### Limitations

- DOCX patches affect the first section only.
- `clearCoreMetadata` does not remove comment authors, revision authors,
  or every possible identity-bearing property.
- The `trackedChanges` inspection count covers inserted/deleted revision
  markers in common content parts; it is not a count of every possible
  WordprocessingML revision type.
- Edited XML parts are reserialized and may differ lexically from their
  original representation.
- PDF writing and patching are not supported.
- XLSX support does not extract cell contents, calculate formulas, or
  manipulate charts, pivots, macros, comments, named ranges, tables,
  conditional formatting, or external data.
