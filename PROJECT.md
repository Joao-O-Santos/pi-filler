# Project

## Objective

Build `pi-filler`, a small deterministic document helper for Pi that
fills DOCX and PDF workflow gaps without requiring a full office suite.

## Intended use

The extension should support manuscript and journal-submission work,
including text extraction, page-aware PDF search, selective page
rendering, DOCX generation, formatting inspection, and narrow
submission-oriented Word formatting patches.

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

GitLab is canonical and is the only release authority. GitHub is a mirror
for verification and Pages.

## Current direction

Implement PDF support first, then Pandoc-based DOCX paths, OOXML
inspection, and a deliberately small typed patch surface. Expand only in
response to real document fixtures.
