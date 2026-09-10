# Architecture

`pi-filler` is a small Pi extension for deterministic document handling.
It delegates common document conversion and PDF extraction to mature
command-line tools and keeps Word-specific mutations narrow and
inspectable.

## Model-facing API

The extension exposes one `filler` tool. Its main dimensions are document
format, action, and view rather than a collection of unrelated tools.

The intended surface is:

```text
filler
format: docx | pdf
action: read | search | write | patch
view: text | formatting | image
```

Unsupported combinations are rejected explicitly.

## External tools

The core requires `pandoc`, `pdftotext`, `pdftocairo`, and `pdfgrep`.
LibreOffice may be used later as an optional DOCX-to-PDF renderer.

Node.js process primitives are used directly. A small local adapter may
normalize repeated stdin, output, timeout, and error handling, but the
package should not grow its own process framework.

## PDF

PDF text reading uses `pdftotext -layout`. Search uses `pdfgrep` so
results remain page-aware. Image views render selected pages through
`pdftocairo -png`.

## DOCX

Pandoc handles ordinary text extraction and document generation.
Formatting inspection and targeted mutations operate directly on OOXML
inside the DOCX ZIP package.

Untouched package parts should remain untouched. Patches should create a
new output file by default and report which package parts changed.

## Validation

Writes and patches validate the ZIP package, required Word parts, edited
XML, relationships, referenced media, and requested mutation state.

## Scope boundary

The project is not a general office suite or broad PDF toolkit. Add
features only when real document fixtures require them.
