# Project

## Objective

Build `pi-filler`, a small deterministic document helper for Pi that
fills DOCX and PDF workflow gaps without requiring a full office suite.

## Intended use

The extension supports manuscript and journal-submission work: text
extraction, page-aware PDF search, selective PDF rendering, DOCX
generation, formatting inspection, and narrow Word-format patches.

## Scope

Pandoc handles DOCX conversion. Poppler command-line tools handle PDF
text and rendering. `pdfgrep` handles page-aware PDF search. Targeted
Word formatting changes operate on OOXML inside the DOCX package.

LibreOffice may later provide optional DOCX-to-PDF rendering. PDF
writing and patching are outside the current scope.

## Constraints

Keep the project small, deterministic, inspectable, and suitable for
short-lived agents. Prefer Node.js and mature command-line primitives
over broad wrappers. Track current upstream releases instead of pinning
versions.

GitLab is canonical and is the only release authority. GitHub is a
mirror for verification and Pages.

## Current state

The planned v0.0.0 core is implemented as a release candidate:

- one `filler` tool with explicit DOCX/PDF operation combinations;
- bounded, cancellable, finite-time external command execution;
- PDF text, search, and selected-page rendering;
- DOCX text, search, and validated transactional generation;
- prefix-flexible OOXML inspection and conservative first-section
  patches;
- transactional patching with dry runs and relationship validation;
- deterministic regression tests for the supported surface.

No v0.0.0 tag or release exists yet. Tagging and publication require
explicit instruction.

## Boundaries

`clearCoreMetadata` is deliberately narrower than anonymization. It
clears common core metadata fields but does not promise removal of every
possible author or identity trace.

Untouched DOCX package parts are preserved byte-for-byte at the part
content level. Edited XML parts are reserialized and are not promised to
be lexically identical to their originals.
