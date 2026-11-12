# Project

## Objective

Build `pi-filler`, a small deterministic document helper for Pi that
fills DOCX, PDF, PPTX, and XLSX workflow gaps without requiring a full
office suite.

## Intended use

The extension supports manuscript and journal-submission work: text
extraction, page-aware PDF search, selective PDF rendering, DOCX
generation, formatting inspection, narrow Word-format patches, and
privacy-preserving spreadsheet template filling and formatting.

## Scope

Pandoc handles DOCX and PPTX conversion. Poppler command-line tools
handle PDF text and rendering. `pdfgrep` handles page-aware PDF search.
Targeted Word and presentation patches operate on OOXML inside their
packages. XLSX operations parse workbook OOXML and CSV data locally,
mutate only needed package parts, and return structural metadata rather
than spreadsheet contents.

LibreOffice provides optional DOCX/PPTX-to-PDF rendering for image
reads. PDF writing and patching are outside the current scope.

## Constraints

Keep the project small, deterministic, inspectable, and suitable for
short-lived agents. Prefer Node.js and mature command-line primitives
over broad wrappers. Track current upstream releases instead of pinning
versions.

GitLab is canonical and is the only release authority. GitHub is a
mirror for verification and Pages.

## Current state

The current pre-1.0 source implements:

- one `filler` tool with explicit DOCX/PDF/PPTX/XLSX operation
  combinations;
- bounded, cancellable, finite-time external command execution;
- isolated per-conversion LibreOffice profiles, cleaned on all exits;
- PDF text, search, and selected-page rendering;
- DOCX Pandoc Markdown text/search with accepted, rejected, or all
  revision views, plus validated transactional generation;
- bounded read-only comment inspection with main-document anchors,
  legacy resolution/reply metadata, and pagination;
- prefix-flexible, order- and text-preserving DOCX OOXML inspection and
  conservative first-section patches;
- transactional ordinary patching with dry runs and relationship
  validation;
- optional DOCX tracked replacement via a separately installed
  `docx-toolkit` executable on `PATH`, without automatic installation or
  download;
- deterministic regression tests for the supported surface;
- content-free XLSX layout inspection, mapped CSV-to-template filling,
  primitive automatic typing, formula and merged-range protection,
  narrow style/layout patches, and optional image rendering through
  LibreOffice;
- Pandoc-based PPTX text extraction, GFM generation, slide rendering,
  and narrow transactional OOXML patches;
- package-owned prompt guidance that keeps deterministic XLSX processing
  local when the model does not need spreadsheet contents.

The repository version may be ahead of the npm release shown in the
README badge. Tagging and publication require explicit instruction.

## Boundaries

`clearCoreMetadata` is deliberately narrower than anonymization. It
clears common core metadata fields but does not promise removal of every
possible author or identity trace.

DOCX tracked replacement accepts `tracked_replace` instead of `patches`
with nonempty string `old`, string `new` (empty deletes), nonempty
string `author`, and optional lowercase 64-hex `expected_sha256`. It
requires a separate nonexistent output, rejects `view` and `dry_run`
even when false, and probes protocol capability before use. It computes
a source hash, checks the optional expected hash, passes the computed
hash for freshness checks, and verifies the output hash. Freshness
checks are not a lock. Publication requires hard-link support. Text and
author travel over local stdin, not command-line arguments; pipes,
memory, and output still require protection. Results omit requested
text, author, and toolkit diagnostics. A subprocess, transport, or
protocol failure can leave an unknown output; inspect it before
retrying. There is no automatic retry or cleanup, and timeout or
cancellation terminates only the direct child.

Replacement is limited to a unique literal single-run match in a simple
main-document paragraph or table. Basic existing revisions and
unthreaded comments outside the target paragraph may be retained only
under strict checks and without an enclosing comment range. Threaded
comments, complex revisions, fields, other stories, cross-run matches,
signed packages, and other unsupported structures are rejected. There is
no accept/reject or comment editing. Structural validation does not
establish Microsoft Word behavior or visual fidelity; inspect a copy in
the target application.

Untouched DOCX, PPTX, and XLSX package parts are preserved byte-for-byte
at the part content level. Edited XML parts are reserialized and are not
promised to be lexically identical to their originals.

Spreadsheet cell values, CSV fields, formula contents, and comments are
tool-local data. Normal XLSX results and errors expose only structural
or operational metadata. That metadata may include worksheet names,
ranges, dimensions, counts, merge addresses, hidden ranges, changed
parts, and output paths. Bounded layout metadata can reveal limited
layout and occupancy information, but not cell values or CSV contents.
XLSX support does not calculate formulas or provide arbitrary content
extraction or spreadsheet programming. PPTX support does not promise
arbitrary XML editing or automatic layout repair; it preserves
unsupported presentation parts but does not manipulate charts, SmartArt,
animations, transitions, notes, masters, themes, or embedded objects.

Model-facing schema descriptions and prompt guidance should make a valid
format/action/view combination possible in one call, distinguish source
and output path roles, state defaults and units, and make result limits
and the XLSX privacy boundary clear before a call.

## Current direction

Keep package-specific usage guidance with `pi-filler` itself. The
registered tool schema, `promptSnippet`, and `promptGuidelines` are the
integration surface; do not add a separate Pi Sych skill or a registry
that duplicates this package-owned knowledge.

Prefer deterministic local XLSX operations when filling or formatting
can be completed without exposing workbook or CSV contents to the model.
When that transformation does not require inspection, do not read CSV or
cell contents into model context merely to carry it out. Use model
interpretation only when the task actually requires semantic inspection.

## Definition of done

- The schema, prompt guidance, README, and architecture description
  agree with implemented behavior.
- Model-facing descriptions make valid calls and safe follow-up calls
  possible without inventing unsupported operations or output semantics.
- Regression coverage protects the key schema and prompt commitments,
  including PPTX conversion, targeted OOXML patches, rendering, and
  content-free XLSX mapping, layout inspection, and rendering;
  project-wide line and function coverage remain above 90%.
- Durable state, generated site files, and synchronization fingerprints
  reflect the accepted documentation; verification passes.

## Previous action

Fixed DOCX patch corruption of order, whitespace, and numeric-looking
text; preserved namespace bindings and section-property order. Switched
DOCX conversion to native Pandoc Markdown with explicit revision views,
isolated LibreOffice profiles, and added bounded read-only comment
inspection. Added preservation, real-Pandoc round-trip, profile cleanup,
review pagination, and malformed-metadata regression tests. Scope stays
narrow: no general Word editing engine, native acceptance/rejection, or
complete modern comment-thread support.

Repaired current-upstream schema typing with native TypeBox enums and
compile-time/runtime regression checks. Kept Markdown formatting with
the repository formatter instead of Biome's newly enabled formatter.
Added public DOCX/PPTX routing and failed-conversion preservation tests,
retaining all existing tests and restoring line coverage. Enabled the
existing GitHub Pages workflow on the mirror; Pages had been disabled.
Added bounded retries for transient Pandoc release-request failures and
private installer temporary-file cleanup, with offline subprocess tests
for recovery, exhausted attempts, client/rate-limit errors, and cleanup.

## Verification

The optional tracked replacement integration passed `make verify` with
its real toolkit executable available: 109 tests, no skips; line,
function, and branch coverage were 97.11%, 97.32%, and 84.57%.
`make site` also passed. The real-backend test uses a synthetic DOCX; it
does not establish Word compatibility. No new tag or release was
created.

## Previous verification

At the previous verification, 82 tests passed with no skips, including
real Pandoc conversion; line/function/branch coverage is
96.65%/97.20%/83.08%. Site build also passed. Each intermediate atomic
tree was independently checked with `make verify`. Git signing remains
enabled.
