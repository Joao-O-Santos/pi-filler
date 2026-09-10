# pi-filler

Deterministic DOCX and PDF tooling for Pi.

`pi-filler` fills the document-handling gaps between ordinary Markdown
workflows and final office-format submission requirements. It delegates
what mature tools already do well and keeps Word-specific mutations
small, typed, and inspectable.

The intended core uses Pandoc for DOCX conversion, Poppler tools for PDF
text and rendering, and `pdfgrep` for page-aware PDF search. Targeted
DOCX formatting changes operate directly on OOXML rather than relying on
a full office suite.

## Status

The repository is scaffolded from `pi-tin`. Implementation has not begun.
See `PLAN.md` for the current build sequence and accepted constraints.

## Principles

-   One small model-facing `filler` tool.
-   Required external tools: Pandoc, `pdftotext`, `pdftocairo`, and
    `pdfgrep`.
-   LibreOffice optional for rendering only.
-   Source files remain untouched by default.
-   Typed, narrow DOCX patches rather than arbitrary XML editing.
-   No broad PDF toolkit or office-suite abstraction.
-   No shared Bakery runtime-utils dependency unless repeated real needs
    justify one later.
-   Current upstream releases are preferred over pinned versions.

## Development

```sh
npm install
make verify
make site
```

GitLab is canonical and is the only release authority. GitHub may mirror
the repository, verify it, and publish GitHub Pages.
