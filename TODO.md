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

- [ ] Implement Pandoc-based DOCX text reads and in-process search.
- [ ] Implement Markdown-to-DOCX writes with optional `reference.docx`.
- [ ] Add deterministic tests using a fake Pandoc executable.

## OOXML

- [ ] Select one ZIP and one namespace-capable XML dependency after a
  small implementation bakeoff.
- [ ] Inspect page setup, margins, sections, numbering, comments,
  tracked changes, and core metadata.
- [ ] Patch page setup, margins, line numbering, page numbering, and
  core metadata/anonymization.
- [ ] Support dry runs, sibling temporary writes, post-write validation,
  and changed/preserved package-part reporting.
- [ ] Add deterministic DOCX fixtures and preservation regression tests.

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
- `npm run typecheck`
- `npm run lint`
- `npm test` (5 passing tests)
- `make verify`
