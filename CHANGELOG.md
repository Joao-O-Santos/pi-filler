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

### Changed

- Search DOCX against the full extraction before truncating results.
- Use Node.js `execFile` for bounded command execution, cancellation,
  and finite timeouts.
- Terminate timed-out external commands even when they ignore
  `SIGTERM`.
- Honor PDF page ranges during search and use pdfgrep's native
  fixed-string mode for literal queries.
- Reject unsupported PDF search views instead of silently accepting
  them.
- Standardize line-number restart modes on OOXML values.
- Rename the misleading anonymization option to `clearCoreMetadata`.
- Make Word-attribute handling tolerant of alternate namespace prefixes.
- Remove stale PDF render outputs and sort generated pages numerically.
- Follow current releases for runtime dependencies.
- Check Markdown prose width directly instead of requiring a Pandoc
  round-trip to be byte-identical.

### Limitations

- DOCX patches affect the first section only.
- `clearCoreMetadata` does not remove comment authors, revision authors,
  or every possible identity-bearing property.
- Edited XML parts are reserialized and may differ lexically from their
  original representation.
- PDF writing and patching are not supported.
- LibreOffice-based DOCX rendering remains future work.
