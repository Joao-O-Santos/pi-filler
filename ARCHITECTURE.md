# Architecture

`pi-filler` is a small Pi extension for deterministic document handling.
It delegates common document conversion and PDF extraction to mature
command-line tools and keeps Word-, presentation-, and
spreadsheet-specific mutations narrow and inspectable.

## Model-facing API

The extension exposes one `filler` tool with explicit `format`,
`action`, and operation-dependent `view` fields:

``` text
format: docx | pdf | pptx | xlsx
action: read | search | write | patch
view: text | formatting | image | structure
```

DOCX supports text, formatting, and image reads; text search;
Markdown-to-DOCX writes; and narrow OOXML patches. PPTX supports text,
formatting, and image reads; text search; GFM-to-PPTX writes; and narrow
OOXML patches. PDF supports text and image reads plus text search. XLSX
supports content-free structure and image reads, CSV-to-template writes,
and formatting/layout patches.

The schema describes operation-specific path roles and defaults.
`format` selects the document or target format, rather than every input:
DOCX and PPTX writes take Markdown at `path` (with an optional
format-specific reference document), and XLSX write takes the workbook
template at `path` plus values from `source_csv`. Other operations use
`path` as the input document. Relative paths resolve from Pi's working
directory; a leading `@` is removed before resolution. Image `output` is
an existing directory or PNG prefix that creates page images, not one
final PNG.

The schema labels page fields as 1-based and inclusive, and XLSX cells
and ranges as A1 notation. `start_cell` supplies the write row and the
default upper-left column; explicit XLSX map destinations are absolute.
Unsupported combinations and operation-irrelevant parameters are
rejected explicitly. Paths and required query/output values cannot be
empty. Writes and patches require an output distinct from the input.
`dry_run` is supported for DOCX and PPTX patches and XLSX writes and
patches; it still requires an output path but does not create or replace
it.

Prompt guidance supplies the compact operation matrix needed to choose a
valid call without retrying through invalid combinations. It names path
roles, required XLSX write and patch fields, PPTX patch fields, XLSX's
content-free privacy boundary, PDF search and XLSX write defaults,
page-range rules, and operation-appropriate narrowing of truncated
results.

That guidance is package-owned. The registered `promptSnippet` and
`promptGuidelines` travel with `pi-filler` and are the appropriate place
for stable usage principles that are not mechanical schema details. In
particular, they make local XLSX processing and non-return of workbook
and CSV contents salient to the model. When a deterministic XLSX write
or patch does not require inspection, the guidance also tells the model
not to read CSV or cell contents into context merely to perform it. Pi
Sych does not need a `pi-filler` skill, package registry entry, or
duplicated description.

## Processes

External commands run through a small local `execFile` adapter. Node.js
provides cancellation, finite timeouts, and maximum buffered output.
Timed-out commands are force-terminated. The adapter otherwise only
normalizes errors and missing-executable messages.

Required executables are `pandoc`, `pdftotext`, `pdftocairo`, and
`pdfgrep`. LibreOffice is optional at runtime and used for DOCX, PPTX,
and XLSX image rendering.

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

XLSX cell values, CSV fields, formula contents, and comments are
tool-local and are never included in normal model-facing results or
errors. Structural reads may report sheet names, active sheet, used
ranges, dimensions, counts for formulas, merges, and styled cells, plus
bounded merged-range addresses and hidden row and column ranges. This
layout metadata can reveal limited layout and occupancy information, but
not cell values or CSV contents.

`fflate` preserves package members and `fast-xml-parser` mutates only
the selected worksheet and, for style patches, `xl/styles.xml`.
Worksheet names are resolved through `xl/workbook.xml` relationships
rather than filename assumptions. `csv-parse` handles source CSV syntax.
Writes use inline strings by default, may infer only empty, numeric, and
boolean primitives in automatic mode, and never infer dates or formulas.
An optional column map maps either all 1-based source ordinals or all
exact, unique CSV headers to absolute Excel columns; maps may omit
source columns but cannot repeat a destination. Destination formulas and
merged ranges cause the entire write to fail.

Style patches cover bold, italic, font size in points, RGB fill,
alignment, wrapping, Excel number-format codes, and a uniform border.
Equivalent style components and cell formats are reused. Column width
uses Excel character-width units; row height uses points. Both apply to
the addressed range's columns and rows. Existing cell values and
formulas remain unchanged during formatting.

XLSX image reads convert workbook print pages to a temporary PDF through
LibreOffice, then use the common PDF renderer to create numbered PNG
files. They do not select a sheet; `first_page` and `last_page` may
select resulting print pages. The temporary PDF is removed and is not an
exported artifact.

XLSX writes and patches require a distinct output path. They run through
the file mutation queue, build and validate a temporary sibling package,
and rename it only after successful validation. Dry runs parse all local
inputs but return only structural operation metadata.

## PPTX conversion and rendering

Pandoc converts PPTX to GFM Markdown for reads and searches. GFM
Markdown converts to PPTX for writes; an optional reference PPTX
supplies the presentation template. PPTX extraction is semantic and
lossy rather than a round-trip representation of slide layout,
animation, themes, or every object.

PPTX image reads convert the presentation to PDF through LibreOffice,
then use the common PDF renderer to create numbered PNG slide images.
Page ranges are 1-based and inclusive.

## PPTX OOXML

PPTX inspection reports slide dimensions in EMUs, slide count, and
common core metadata. Patches support slide dimensions, common core
metadata, and text replacement within slide text paragraphs.
Replacements can target one slide or all slides and can span DrawingML
text runs; replacement text inherits the first matched run's formatting.
Edited XML parts are reserialized, while untouched package parts are
preserved at the part content level.

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

PPTX support does not promise arbitrary XML editing, preservation of
every PowerPoint feature, or automatic layout repair. It does not
manipulate charts, SmartArt, animations, transitions, notes, masters,
themes, or embedded objects except by preserving untouched package
parts.

The project is not a general office suite, broad PDF toolkit, complete
DOCX anonymizer, or spreadsheet programming environment. XLSX support
does not extract arbitrary content, evaluate or generate formulas, or
manipulate charts, pivots, macros, comments, named ranges, data
validation, tables, external data, or conditional formatting. Existing
unsupported package parts are preserved where possible. New OOXML
features should be driven by real document fixtures rather than
speculative coverage.
