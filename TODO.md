# Implementation TODO

This checklist tracks the temporary work needed for `v0.0.0`. Workers
mark completed items and record verification beneath their assigned
slice. The supervisor reviews and commits each coherent increment.

## Foundation

- [x] Define the Google-compatible `filler` schema and runtime operation
  validation.
- [x] Add bounded, abort-aware process execution and capability errors.
- [x] Register the Pi tool with path normalization, output truncation,
  and file-mutation queue integration.
- [x] Add focused unit tests for schema dispatch and process failures.

## PDF

- [x] Implement layout-preserving PDF text reads with optional page
  ranges.
- [x] Implement page-aware literal and regular-expression search.
- [x] Render only requested PDF pages to PNG output files.
- [x] Add deterministic tests using fake command executables.

## DOCX text paths

- [x] Implement Pandoc-based DOCX text reads and in-process search.
- [x] Implement Markdown-to-DOCX writes with optional `reference.docx`.
- [x] Add deterministic tests using a fake Pandoc executable.

## OOXML

- [x] Select `fflate` for ZIP preservation and `fast-xml-parser` for
  namespace-aware XML inspection and mutation.
- [x] Inspect page setup, margins, sections, numbering, comments,
  tracked changes, and core metadata.
- [x] Patch page setup, margins, line numbering, page numbering, and
  core metadata/anonymization.
- [x] Support dry runs, sibling temporary writes, post-write validation,
  and changed/preserved package-part reporting.
- [x] Add deterministic DOCX fixtures and preservation regression tests.

## Release preparation

- [ ] Update durable architecture, project status, README, and
  changelog.
- [ ] Remove `PLAN.md` and this file after all accepted decisions are
  represented durably.
- [ ] Run `make verify` and `make site`.
- [ ] Obtain an independent implementation and release-readiness review.
- [ ] Push `main`, confirm GitLab CI passes, then tag and push `v0.0.0`.

## Verification

- `npm run format:fix`
- `make verify` (10 passing tests)
- `git diff --check`
