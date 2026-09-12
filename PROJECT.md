# Project

## Objective

Build `pi-filler`, a small deterministic document helper for Pi that
fills DOCX, PDF, and XLSX workflow gaps without requiring a full office
suite.

## Intended use

The extension supports manuscript and journal-submission work: text
extraction, page-aware PDF search, selective PDF rendering, DOCX
generation, formatting inspection, narrow Word-format patches, and
privacy-preserving spreadsheet template filling and formatting.

## Scope

Pandoc handles DOCX conversion. Poppler command-line tools handle PDF
text and rendering. `pdfgrep` handles page-aware PDF search. Targeted
Word formatting changes operate on OOXML inside the DOCX package. XLSX
operations parse workbook OOXML and CSV data locally, mutate only needed
package parts, and return structural metadata rather than spreadsheet
contents.

LibreOffice provides optional DOCX-to-PDF rendering for image reads. PDF
writing and patching are outside the current scope.

## Constraints

Keep the project small, deterministic, inspectable, and suitable for
short-lived agents. Prefer Node.js and mature command-line primitives
over broad wrappers. Track current upstream releases instead of pinning
versions.

GitLab is canonical and is the only release authority. GitHub is a
mirror for verification and Pages.

## Current state

The current pre-1.0 source implements:

- one `filler` tool with explicit DOCX/PDF/XLSX operation combinations;
- bounded, cancellable, finite-time external command execution;
- PDF text, search, and selected-page rendering;
- DOCX text, search, and validated transactional generation;
- prefix-flexible OOXML inspection and conservative first-section
  patches;
- transactional patching with dry runs and relationship validation;
- deterministic regression tests for the supported surface;
- structural-only XLSX inspection, CSV-to-template filling, primitive
  automatic typing, formula protection, and narrow style/layout patches.

The repository version may be ahead of the npm release shown in the
README badge. Tagging and publication require explicit instruction.

## Boundaries

`clearCoreMetadata` is deliberately narrower than anonymization. It
clears common core metadata fields but does not promise removal of every
possible author or identity trace.

Untouched DOCX and XLSX package parts are preserved byte-for-byte at the
part content level. Edited XML parts are reserialized and are not
promised to be lexically identical to their originals.

Spreadsheet cells, CSV fields, formulas, comments, and other workbook
text are tool-local data. Normal XLSX results and errors expose only
structural or operational metadata. XLSX support does not calculate
formulas or provide arbitrary content extraction or spreadsheet
programming.

Model-facing schema descriptions and prompt guidance should make a valid
format/action/view combination possible in one call, distinguish source
and output path roles, and state defaults and units.
