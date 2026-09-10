# Changelog

## v0.0.0 (released)

**Features:** - PDF text extraction with layout preservation and
optional page ranges - PDF page-aware search (literal and regex) with
page numbers in results - PDF page rendering to PNG files - DOCX text
extraction and search through Pandoc - DOCX generation from Markdown
with optional reference document - DOCX formatting inspection: page
setup, margins, sections, numbering, comments, tracked changes, core
metadata - DOCX formatting patches: page size/orientation, margins,
line/page numbering, core metadata, anonymization - Dry-run support for
patches - Transactional writes with temporary sibling files and atomic
renames - Bounded output (50KB, 2000 lines) - File mutation queue
integration with Pi - Abort signal propagation and finite timeouts -
Deterministic test suite with fake command executables

**Implementation notes:** - Uses `fflate` for ZIP preservation and
`fast-xml-parser` for OOXML - Pandoc handles DOCX conversions; Poppler
tools handle PDF operations - OOXML namespace handling preserves
existing prefixes during round-trips - All unrelated package parts
remain content-identical in patches - Line numbering modes (`off`,
`continuous`, `newPage`, `newSection`) map to OOXML semantics - Core
metadata respects namespace distinctions (dc:, cp:) - Anonymization
clears configurable metadata fields

**Known limitations:** - Patches affect the first section only
(multi-section documents are read-only for formatting changes) - OOXML
namespace handling may reorder elements and lose non-essential lexical
constructs - Content-Types and all relationship targets are validated;
malformed OPC packages are rejected - PDF write and patch operations are
not yet supported - LibreOffice rendering is deferred

**Testing:** - 10 deterministic tests with fake executables - Coverage:
schema validation, path normalization, PDF/DOCX operations, OOXML
fixtures, preservation guarantees, error handling - All checks pass:
typecheck, lint, format, markdown, npm pack
