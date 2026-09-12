# Architecture

`pi-filler` is a small Pi extension for deterministic document handling.
It delegates common document conversion and PDF extraction to mature
command-line tools and keeps Word- and spreadsheet-specific mutations
narrow and inspectable.

## Model-facing API

The extension exposes one `filler` tool with explicit `format`,
`action`, and operation-dependent `view` fields:

``` text
format: docx | pdf | xlsx
action: read | search | write | patch
view: text | formatting | image | structure
```

DOCX supports text, formatting, and image reads; text search;
Markdown-to-DOCX writes; and narrow OOXML patches. PDF supports text and
image reads plus text search. XLSX supports structure-only reads,
CSV-to-template writes, and formatting/layout patches.

The schema describes operation-specific path roles and defaults. For
DOCX write, `path` is Markdown source. For XLSX write, `path` is the
workbook template and `source_csv` supplies values. Other operations use
`path` as the input document. Relative paths resolve from Pi's working
directory; a leading `@` is removed before resolution.

Unsupported combinations and operation-irrelevant parameters are
rejected explicitly. Paths and required query/output values cannot be
empty. Writes and patches require an output distinct from the input.
`dry_run` is supported for DOCX patches and XLSX writes and patches.

Prompt guidance supplies the compact operation matrix needed to choose a
valid call without retrying through invalid combinations.

## Processes

External commands run through a small local `execFile` adapter. Node.js
provides cancellation, finite timeouts, and maximum buffered output.
Timed-out commands are force-terminated. The adapter otherwise only
normalizes errors and missing-executable messages.

Required executables are `pandoc`, `pdftotext`, `pdftocairo`, and
`pdfgrep`. LibreOffice is optional at runtime and used for DOCX image
rendering.

## PDF

PDF text reading uses `pdftotext -layout`. Search uses `pdfgrep` so
matches remain page-aware and may be restricted to a page range. Literal
search uses pdfgrep's fixed-string mode. Malformed search output is
rejected rather than converted into synthetic page numbers.

Search text is bounded before it reaches the model, and tool details
retain only the match count and truncation metadata.

Image reads render selected pages through `pdftocairo -png`. Rendering
happens under a temporary sibling directory. Existing outputs for the
requested prefix are replaced only after `pdftocairo` succeeds, so a
failed render does not destroy a previous successful result. Generated
page files are sorted by numeric page number.

## XLSX

XLSX data is tool-local. Workbook cells, CSV fields, formulas, comments,
and other spreadsheet text are never included in normal model-facing
results or errors. Structural reads report sheet names, active sheet,
used ranges, dimensions, and counts for formulas, merges, and styled
cells.

`fflate` preserves package members and `fast-xml-parser` mutates only
the selected worksheet and, for style patches, `xl/styles.xml`.
Worksheet names are resolved through `xl/workbook.xml` relationships
rather than filename assumptions. `csv-parse` handles source CSV syntax.
Writes use inline strings by default, may infer only empty, numeric, and
boolean primitives in automatic mode, and never infer dates or formulas.
Destination formulas cause the entire write to fail.

Style patches cover bold, italic, font size in points, RGB fill,
alignment, wrapping, Excel number-format codes, and a uniform border.
Equivalent style components and cell formats are reused. Column width
uses Excel character-width units; row height uses points. Both apply to
the addressed range's columns and rows. Existing cell values and
formulas remain unchanged during formatting.

XLSX writes and patches require a distinct output path. They run through
the file mutation queue, build and validate a temporary sibling package,
and rename it only after successful validation. Dry runs parse all local
inputs but return only structural operation metadata.

## DOCX text and generation

Pandoc converts DOCX to Markdown for reads and searches. Search examines
the full bounded Pandoc extraction, then returns bounded match text plus
a match count instead of duplicating the full match set in tool details.

DOCX generation writes to a temporary sibling, validates the generated
ZIP/OOXML package, and renames it to the requested output only after
validation succeeds.

## DOCX OOXML

`fflate` reads and writes ZIP packages. `fast-xml-parser` handles the
small set of XML parts needed by the supported patches. XML is checked
for well-formedness before the parsed representation is accepted.

Inspection and mutation match XML names by local name and preserve or
derive the document's namespace prefix when writing Word attributes.
This supports ordinary alternate-prefix documents but is not a general
namespace-normalizing OOXML engine.

Page dimensions and margins use OOXML twips (twentieths of a point). The
current patch surface covers:

- first-section page size and orientation;
- first-section margins;
- line numbering modes `off`, `continuous`, `newPage`, and `newSection`;
- page-number start and format;
- common core metadata;
- `clearCoreMetadata`, which clears common core metadata only.

Orientation-only changes keep page dimensions consistent. Untouched ZIP
package parts remain content-identical. Edited XML parts are
reserialized, so lexical XML details inside those parts may change.

## Validation

DOCX inspection and mutations validate required package parts and parse
every XML and relationship part. XLSX operations validate workbook
metadata, worksheets, styles when used, and all package relationships.
Internal targets from every `.rels` part must resolve to an existing
package part.

Patches reject empty requests and incompatible line-number settings.
Patches are validated against the requested resulting formatting. Writes
and patches use temporary sibling files and atomic renames. Source files
are not modified.

## Scope boundary

The project is not a general office suite, broad PDF toolkit, complete
DOCX anonymizer, or spreadsheet programming environment. XLSX support
does not extract arbitrary content, evaluate or generate formulas, or
manipulate charts, pivots, macros, comments, named ranges, data
validation, tables, external data, or conditional formatting. Existing
unsupported package parts are preserved where possible. New OOXML
features should be driven by real document fixtures rather than
speculative coverage.
