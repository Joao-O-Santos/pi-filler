# Project

## Objective

Build `pi-filler`, a small deterministic document helper for Pi that
fills DOCX and PDF workflow gaps without requiring a full office suite.

## Intended use

The extension supports manuscript and journal-submission work, including
text extraction, page-aware PDF search, selective page rendering, DOCX
generation, formatting inspection, and narrow submission-oriented Word
formatting patches.

## Scope

The core supports PDF and DOCX. Pandoc handles conversions it already
does well. Poppler command-line tools handle PDF text and rendering.
`pdfgrep` provides page-aware PDF search. Word-specific formatting is
handled through constrained OOXML inspection and patching.

LibreOffice may later provide optional DOCX-to-PDF rendering. It is not
required for the core and should not normalize ordinary DOCX output.

## Constraints

Keep the project small, deterministic, inspectable, and suitable for
short-lived agents. Prefer mature command-line tools and Node.js
primitives over broad wrapper libraries. Track current upstream releases
instead of pinning versions.

GitLab is canonical and is the only release authority. GitHub is a
mirror for verification and Pages.

## Status

The extension is complete and tested for v0.0.0. All planned Foundation,
PDF, DOCX, and OOXML slices are implemented and verified. The tool
integrates with Pi's file mutation queue, respects abort signals and
timeouts, bounds output, and handles errors clearly.

## Definition of done

- [x] Extension builds, typechecks, lints, and tests cleanly
- [x] All tests pass deterministically
- [x] File mutation queue integration works for DOCX writes/patches
- [x] Output is bounded and truncation is reported
- [x] Abort signals flow through operations
- [x] Errors include actionable messages
- [x] Path normalization handles `@` and relative paths
- [x] DOCX inspection and patching preserve untouched parts
- [x] Dry-run support works correctly
- [x] README documents the tool and constraints
- [x] CHANGELOG describes the v0.0.0 release
- [x] All durable documentation reflects the implementation

## Immediate next step

Push to main, observe CI/CD, and prepare for npm publication.
