# Implementation plan

## Goal

Build `pi-filler` as a small deterministic Pi extension for reading,
searching, generating, inspecting, and narrowly patching DOCX and PDF
files. It should cover the boring office-document gaps left by Pandoc
without becoming a general office suite.

Target roughly 450--600 production TypeScript lines for the first useful
release, excluding tests and optional Lua filters.

## Constraints

-   Start from the `pi-tin` conventions and keep them intact.
-   Track current upstream releases rather than pinning versions.
-   Use Node.js process primitives directly where practical.
-   Do not create a shared Bakery runtime-utils dependency.
-   Require `pandoc`, `pdftotext`, `pdftocairo`, and `pdfgrep`.
-   Keep LibreOffice optional and use it only for rendering/conversion.
-   Keep the model-facing API to one `filler` tool.
-   Preserve source files by default and make mutations transactional.
-   Expand OOXML support only when fixtures demonstrate a need.

## 1. Define the tool surface

Implement one model-facing tool with the conceptual shape:

```text
filler
format: docx | pdf
action: read | search | write | patch
view: text | formatting | image
```

Use TypeBox for a narrow discriminated schema. Reject unsupported
combinations rather than accepting arbitrary option bags.

Initial combinations:

-   DOCX `read/text`
-   DOCX `read/formatting`
-   DOCX `search`
-   DOCX `write`
-   DOCX `patch`
-   PDF `read/text`
-   PDF `read/image`
-   PDF `search`

Defer PDF writing and patching.

## 2. Add capability detection

Check required external executables and return concise actionable errors
when they are absent.

Required:

```text
pandoc
pdftotext
pdftocairo
pdfgrep
```

Optional later:

```text
libreoffice
```

Do not build a dependency-management subsystem.

## 3. Add minimal process handling

Use `node:child_process` directly. Prefer `execFile` for bounded buffered
commands and `spawn` only when streaming materially helps.

If repeated call sites need the same stdin, timeout, Buffer output, or
error normalization, add one small local adapter. Do not copy the old
large `pi-pew-pew` process wrapper unless real requirements prove the
Node primitives insufficient.

## 4. Implement PDF first

Keep `src/pdf.ts` small and mostly concerned with argument construction
and result parsing.

Text:

```text
pdftotext -layout input.pdf -
```

Search:

```text
pdfgrep <whitelisted options> query input.pdf
```

Return page-aware matches so the caller can escalate naturally from
search to text reading to rendering.

Image:

```text
pdftocairo -png <selected page range> input.pdf output-prefix
```

Render only requested pages. Do not add a broad PDF manipulation layer.

## 5. Implement DOCX text paths

Add `src/pandoc.ts` for narrow Pandoc invocation helpers.

Read:

```text
DOCX -> Pandoc -> Markdown
```

Search:

```text
DOCX -> Pandoc -> Markdown -> small in-process JS search
```

Write:

```text
Markdown -> Pandoc + reference.docx + optional CSL/bibliography/Lua
         -> DOCX
```

Only expose Pandoc options needed by real manuscript workflows.

## 6. Choose ZIP and XML dependencies

Run a small fixture bakeoff before committing to libraries.

Prefer exactly:

-   one small mature ZIP package;
-   one small namespace-capable XML package.

Candidates include `fflate` or `JSZip` for ZIP handling and a small XML
library capable of namespace-aware mutation.

Selection criteria are preservation behavior, dependency size, API
clarity, and how little glue code is required. Do not choose a full DOCX
object model merely because one exists.

## 7. Add DOCX formatting inspection

Inspect relevant OOXML parts and return structured state rather than raw
XML.

Initial report should cover, where present:

-   page size and orientation;
-   margins;
-   section count;
-   line numbering;
-   page numbering;
-   selected named styles;
-   comments;
-   tracked changes;
-   document metadata.

Raw XML may exist only as an internal/debug escape hatch.

## 8. Add narrow OOXML patching

Implement preservation-first typed mutations for the first real journal
submission needs:

-   page size and orientation;
-   margins;
-   continuous/restart/off line numbering;
-   page numbering start and format;
-   named paragraph/run style properties;
-   document metadata and anonymization;
-   selected `settings.xml` properties.

Only add headers, footers, section-break peculiarities, fields, or
numbering-specific behavior when fixtures require them.

Patching semantics:

1.  Keep the source immutable by default.
2.  Copy the DOCX package.
3.  Mutate only targeted XML parts.
4.  Validate the result.
5.  Write a new output DOCX.
6.  Report changed and preserved package parts.

Support `dry_run: true` early.

## 9. Validate every write and patch

Automatically verify:

-   valid ZIP package;
-   `[Content_Types].xml` exists;
-   required Word parts exist;
-   edited XML parses;
-   relationships resolve;
-   referenced media exists;
-   requested formatting state is present.

Tests should distinguish content preservation from binary ZIP identity:
a no-op rewrite may change ZIP serialization while unrelated package
parts should remain content-identical.

## 10. Build regression fixtures

Before expanding the patch surface, add representative fixtures for:

-   ordinary Pandoc manuscript;
-   real journal `reference.docx`;
-   custom styles;
-   multiple sections;
-   headers and footers;
-   continuous line numbering;
-   comments and tracked changes;
-   fields and cross-references;
-   figures, tables, and captions;
-   footnotes;
-   unknown/custom OOXML parts.

For every new mutation, test that only expected XML parts change.

## 11. Add optional visual DOCX inspection last

If LibreOffice is available:

```text
DOCX -> LibreOffice headless -> PDF -> pdftocairo -> selected PNG pages
```

This is an optional rendering path, not a normalization step. Do not
round-trip DOCX through LibreOffice during ordinary writes or patches.

## 12. Integrate with Pi Sych after the core is stable

Keep `pi-filler` standalone. Pi Sych should call it from writing,
reviewing, or future automation-style workflows rather than absorbing
its document mechanics into Pi Sych.

Do not add a public Pi Sych office skill merely to expose this package if
existing skills can invoke it naturally.

## Suggested source layout

Start flat:

```text
src/
├── index.ts
├── process.ts   # only if repeated Node calls justify it
├── pandoc.ts
├── pdf.ts
└── docx.ts
```

Do not pre-split `docx.ts` into package/inspect/patch/validate modules.
Split only when the file becomes materially harder to navigate.

## Implementation order

```text
Pi schema and capability checks
-> PDF
-> Pandoc DOCX read/search/write
-> OOXML inspection
-> minimal typed patches
-> validation and fixture expansion
-> optional rendering
-> Pi Sych integration
```

## Completion criteria for v0.1

A useful first release should:

-   read and search PDFs with page-aware results;
-   render selected PDF pages;
-   read and search DOCX through Pandoc;
-   generate ordinary DOCX with reference-document support;
-   inspect the main submission-relevant Word formatting state;
-   apply the initial typed patches transactionally;
-   preserve unrelated OOXML parts in fixtures;
-   pass `make verify` and build both Pages sites;
-   remain small enough to understand in one sitting.

After these decisions are implemented and documented in durable project
files, remove `PLAN.md`.
