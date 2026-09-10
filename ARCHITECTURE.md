# Architecture

`pi-filler` is a small Pi extension for deterministic document handling.
It delegates common document conversion and PDF extraction to mature
command-line tools and keeps Word-specific mutations narrow and
inspectable.

## Model-facing API

The extension exposes one `filler` tool:

```text
format: docx | pdf
action: read | search | write | patch
view: text | formatting | image
```

Unsupported combinations are rejected explicitly.

## Processes

External commands run through a small local `execFile` adapter. Node.js
provides cancellation, finite timeouts, and maximum buffered output.
The adapter only normalizes errors and missing-executable messages.

Required executables are `pandoc`, `pdftotext`, `pdftocairo`, and
`pdfgrep`. LibreOffice is optional future work for DOCX rendering.

## PDF

PDF text reading uses `pdftotext -layout`. Search uses `pdfgrep` so
matches remain page-aware. Image reads render selected pages through
`pdftocairo -png`.

Before rendering, matching old output files are removed. Generated page
files are discovered after the command and sorted by numeric page
number.

## DOCX text and generation

Pandoc converts DOCX to Markdown for reads and searches. Search examines
the full Pandoc extraction and truncates only returned match output.

DOCX generation writes to a temporary sibling, validates the generated
ZIP/OOXML package, and renames it to the requested output only after
validation succeeds.

## DOCX OOXML

`fflate` reads and writes ZIP packages. `fast-xml-parser` handles the
small set of XML parts needed by the supported patches.

Inspection and mutation match XML names by local name and preserve or
derive the document's namespace prefix when writing Word attributes.
This supports ordinary alternate-prefix documents but is not a general
namespace-normalizing OOXML engine.

The current patch surface covers:

- first-section page size and orientation;
- first-section margins;
- line numbering modes `off`, `continuous`, `newPage`, and
  `newSection`;
- page-number start and format;
- common core metadata;
- `clearCoreMetadata`, which clears common core metadata only.

Untouched ZIP package parts remain content-identical. Edited XML parts
are reserialized, so lexical XML details inside those parts may change.

## Validation

Inspection and mutations validate required package parts and parse every
XML and relationship part. Internal targets from every `.rels` part
must resolve to an existing package part.

Patches are validated against the requested resulting formatting. Writes
and patches use temporary sibling files and atomic renames. Source files
are not modified.

## Scope boundary

The project is not a general office suite, broad PDF toolkit, or
complete DOCX anonymizer. New OOXML features should be driven by real
document fixtures rather than speculative coverage.
