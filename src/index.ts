import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, extname, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import {
	DEFAULT_MAX_BYTES,
	truncateHead,
	withFileMutationQueue,
} from "@earendil-works/pi-coding-agent";
import { type Static, Type } from "typebox";
import { type DocxPatchSet, inspectDocx, inspectDocxReview, patchDocx } from "./docx.js";
import {
	readDocxText,
	readPptxText,
	searchDocx,
	searchPptx,
	writeDocx,
	writePptx,
} from "./pandoc.js";
import { readPdfText, renderPdfPages, searchPdf } from "./pdf.js";
import { annotatePdf, listPdfComments, type PdfComment } from "./pdf-annotate.js";
import { inspectPptx, type PptxPatchSet, patchPptx } from "./pptx.js";
import { commandError, runCommand } from "./process.js";
import { replaceTracked, validateTrackedReplacement } from "./tracked-replace.js";
import {
	fillXlsxFromCsv,
	inspectXlsxStructure,
	patchXlsx,
	patchXlsxFormula,
	patchXlsxFormulaSheets,
	recalculateXlsx,
	validateXlsxFormula,
	validateXlsxFormulaSheets,
	validateXlsxFormulas,
	type XlsxPatchSet,
} from "./xlsx.js";

const docxPatchesSchema = Type.Object(
	{
		page: Type.Optional(
			Type.Object(
				{
					width: Type.Optional(
						Type.Integer({ description: "Page width in twips (1/20 point)", minimum: 1 }),
					),
					height: Type.Optional(
						Type.Integer({ description: "Page height in twips (1/20 point)", minimum: 1 }),
					),
					orientation: Type.Optional(Type.Enum(["portrait", "landscape"], { type: "string" })),
				},
				{
					additionalProperties: false,
					description: "First-section page size and orientation",
					minProperties: 1,
				},
			),
		),
		margins: Type.Optional(
			Type.Object(
				{
					top: Type.Optional(Type.Integer()),
					bottom: Type.Optional(Type.Integer()),
					left: Type.Optional(Type.Integer()),
					right: Type.Optional(Type.Integer()),
					header: Type.Optional(Type.Integer()),
					footer: Type.Optional(Type.Integer()),
					gutter: Type.Optional(Type.Integer()),
				},
				{
					additionalProperties: false,
					description: "First-section margins in twips (1/20 point)",
					minProperties: 1,
				},
			),
		),
		lineNumbering: Type.Optional(
			Type.Object(
				{
					mode: Type.Optional(
						Type.Enum(["off", "continuous", "newPage", "newSection"], { type: "string" }),
					),
					start: Type.Optional(Type.Integer({ description: "Starting line number", minimum: 0 })),
					count_by: Type.Optional(Type.Integer({ description: "Numbering interval", minimum: 1 })),
				},
				{
					additionalProperties: false,
					description: "First-section line numbering",
					minProperties: 1,
				},
			),
		),
		pageNumbering: Type.Optional(
			Type.Object(
				{
					start: Type.Optional(Type.Integer({ description: "Starting page number", minimum: 0 })),
					format: Type.Optional(
						Type.String({
							description: "Word numbering format, such as decimal or upperRoman",
							minLength: 1,
						}),
					),
				},
				{
					additionalProperties: false,
					description: "First-section page numbering",
					minProperties: 1,
				},
			),
		),
		metadata: Type.Optional(
			Type.Object(
				{
					title: Type.Optional(Type.String()),
					subject: Type.Optional(Type.String()),
					creator: Type.Optional(Type.String()),
					keywords: Type.Optional(Type.String()),
					description: Type.Optional(Type.String()),
					lastModifiedBy: Type.Optional(Type.String()),
				},
				{
					additionalProperties: false,
					description: "Common docProps/core.xml metadata fields",
					minProperties: 1,
				},
			),
		),
		clearCoreMetadata: Type.Optional(
			Type.Boolean({
				description: "Clear common fields in docProps/core.xml; not comment or revision authors",
			}),
		),
	},
	{
		additionalProperties: false,
		description: "DOCX changes for action=patch; first-section settings use OOXML units",
		minProperties: 1,
	},
);

const pptxPatchesSchema = Type.Object(
	{
		slideSize: Type.Optional(
			Type.Object(
				{
					width: Type.Optional(Type.Integer({ description: "Slide width in EMUs", minimum: 1 })),
					height: Type.Optional(Type.Integer({ description: "Slide height in EMUs", minimum: 1 })),
				},
				{
					additionalProperties: false,
					description: "Presentation slide size in EMUs",
					minProperties: 1,
				},
			),
		),
		metadata: Type.Optional(
			Type.Object(
				{
					title: Type.Optional(Type.String()),
					subject: Type.Optional(Type.String()),
					creator: Type.Optional(Type.String()),
					keywords: Type.Optional(Type.String()),
					description: Type.Optional(Type.String()),
					lastModifiedBy: Type.Optional(Type.String()),
				},
				{
					additionalProperties: false,
					description: "Common docProps/core.xml metadata fields",
					minProperties: 1,
				},
			),
		),
		replaceText: Type.Optional(
			Type.Array(
				Type.Object(
					{
						find: Type.String({
							description: "Text to find within a slide text paragraph",
							minLength: 1,
						}),
						replace: Type.String({ description: "Replacement text; may be empty" }),
						slide: Type.Optional(
							Type.Integer({
								description: "1-based slide number; omit for all slides",
								minimum: 1,
							}),
						),
					},
					{ additionalProperties: false, minProperties: 2 },
				),
				{ minItems: 1, description: "Targeted text replacements in slide text" },
			),
		),
		clearCoreMetadata: Type.Optional(
			Type.Boolean({ description: "Clear common core metadata fields" }),
		),
	},
	{
		additionalProperties: false,
		description: "PPTX slide, metadata, and targeted text changes for action=patch",
		minProperties: 1,
	},
);

const xlsxPatchesSchema = Type.Object(
	{
		font: Type.Optional(
			Type.Object(
				{
					bold: Type.Optional(Type.Boolean()),
					italic: Type.Optional(Type.Boolean()),
					size: Type.Optional(
						Type.Number({ description: "Font size in points", exclusiveMinimum: 0 }),
					),
				},
				{ additionalProperties: false, description: "Font changes", minProperties: 1 },
			),
		),
		fill: Type.Optional(
			Type.String({
				description: "Six hexadecimal RGB digits, without # or alpha",
				pattern: "^[0-9A-Fa-f]{6}$",
			}),
		),
		alignment: Type.Optional(
			Type.Object(
				{
					horizontal: Type.Optional(Type.Enum(["left", "center", "right"], { type: "string" })),
					vertical: Type.Optional(Type.Enum(["top", "center", "bottom"], { type: "string" })),
					wrap_text: Type.Optional(Type.Boolean()),
				},
				{ additionalProperties: false, description: "Cell alignment changes", minProperties: 1 },
			),
		),
		number_format: Type.Optional(
			Type.String({ description: "Excel number-format code", minLength: 1 }),
		),
		border: Type.Optional(
			Type.Object(
				{ style: Type.Optional(Type.Enum(["thin", "medium", "thick"], { type: "string" })) },
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		column_width: Type.Optional(
			Type.Number({
				description: "Column width in Excel character-width units",
				exclusiveMinimum: 0,
				maximum: 255,
			}),
		),
		row_height: Type.Optional(
			Type.Number({ description: "Row height in points", exclusiveMinimum: 0, maximum: 409 }),
		),
	},
	{
		additionalProperties: false,
		description: "XLSX formatting and layout changes for action=patch",
		minProperties: 1,
	},
);

const xlsxFormulaSchema = Type.Object(
	{
		cell: Type.String({ description: "Single unmerged A1 cell (case-insensitive)", minLength: 2 }),
		formula: Type.String({
			description:
				"Formula starting with = and a nonblank XML-safe body, no CR; at most 8192 UTF-16 code units; stored, never evaluated",
			minLength: 2,
			maxLength: 8192,
		}),
	},
	{ additionalProperties: false },
);

const pdfCommentSchema = Type.Object(
	{
		page: Type.Integer({ description: "1-based page containing the text", minimum: 1 }),
		text: Type.String({
			description:
				"Exact text to highlight; words separated by spaces, case-sensitive, matching exactly once on the page",
			minLength: 1,
		}),
		comment: Type.String({ description: "Comment text stored in the annotation", minLength: 1 }),
		author: Type.String({ description: "Comment author name", minLength: 1 }),
	},
	{ additionalProperties: false },
);

export const fillerSchema = Type.Object(
	{
		format: Type.Enum(["docx", "pdf", "pptx", "xlsx"], {
			type: "string",
			description: "Document or target format; selects supported operations and path roles",
		}),
		action: Type.Enum(["read", "search", "write", "patch"], {
			type: "string",
			description: "Operation; supported combinations depend on format and view",
		}),
		view: Type.Optional(
			Type.Enum(["text", "formatting", "image", "structure", "review", "comments"], {
				type: "string",
				description: "Read/search representation; omit for write and patch",
			}),
		),
		path: Type.String({
			description:
				"Input path: document for read/search/patch, Markdown source for DOCX/PPTX write, or template for XLSX write",
			minLength: 1,
		}),
		output: Type.Optional(
			Type.String({
				description:
					"New document path, or existing output directory/PNG prefix for view=image; image reads create PNG files",
				minLength: 1,
			}),
		),
		query: Type.Optional(
			Type.String({ description: "DOCX, PDF, or PPTX search query", minLength: 1 }),
		),
		review_offset: Type.Optional(
			Type.Integer({
				minimum: 0,
				default: 0,
				description: "DOCX review read: zero-based comment offset",
			}),
		),
		review_limit: Type.Optional(
			Type.Integer({
				minimum: 1,
				maximum: 50,
				default: 20,
				description: "DOCX review read: maximum comments; byte budget may return fewer",
			}),
		),
		track_changes: Type.Optional(
			Type.Enum(["accept", "reject", "all"], {
				type: "string",
				default: "accept",
				description: "DOCX text read/search revision view; all includes review spans and comments",
			}),
		),
		reference_docx: Type.Optional(
			Type.String({ description: "Reference DOCX for writes", minLength: 1 }),
		),
		patches: Type.Optional(docxPatchesSchema),
		tracked_replace: Type.Optional(
			Type.Object(
				{
					old: Type.String({
						minLength: 1,
						description: "Unique single-run main-story text to replace",
					}),
					new: Type.String({ description: "Replacement text; empty deletes" }),
					author: Type.String({ minLength: 1, description: "Tracked revision author" }),
					expected_sha256: Type.Optional(Type.String({ pattern: "^[0-9a-f]{64}$" })),
				},
				{
					additionalProperties: false,
					description:
						"DOCX patch alternative to patches; requires optional docxquisite on PATH and a distinct output (an existing output is replaced on success; overwrite_original replaces the source). No dry_run; restricted documents only, not general Word editing",
				},
			),
		),
		pptx_patches: Type.Optional(pptxPatchesSchema),
		reference_pptx: Type.Optional(
			Type.String({ description: "Reference PPTX for writes", minLength: 1 }),
		),
		source_csv: Type.Optional(
			Type.String({ description: "CSV source path for XLSX write", minLength: 1 }),
		),
		sheet: Type.Optional(
			Type.String({ description: "Worksheet name for XLSX write or patch", minLength: 1 }),
		),
		start_cell: Type.Optional(
			Type.String({
				description:
					"XLSX start row and default upper-left destination cell in A1 notation, such as A2; column_map destinations are absolute",
				minLength: 1,
			}),
		),
		range: Type.Optional(
			Type.String({
				description: "XLSX target cell or range in A1 notation, such as A2:D10",
				minLength: 1,
			}),
		),
		has_header: Type.Optional(
			Type.Boolean({ default: true, description: "Skip the first CSV record as a header" }),
		),
		value_mode: Type.Optional(
			Type.Enum(["text", "auto"], {
				type: "string",
				default: "text",
				description: "Write CSV fields as text, or infer only empty, number, and boolean values",
			}),
		),
		column_map: Type.Optional(
			Type.Record(Type.String({ minLength: 1 }), Type.String({ minLength: 1 }), {
				description:
					"CSV source columns to absolute XLSX destination columns: all 1-based ordinals or all exact headers to A1 column letters",
				minProperties: 1,
			}),
		),
		pdf_comments: Type.Optional(
			Type.Array(pdfCommentSchema, {
				minItems: 1,
				maxItems: 100,
				description:
					"PDF patch: 1–100 highlight comments on existing text, validated before one atomic output. Each text must match exactly once on its page; rotated pages and signed PDFs are rejected. Results return new comment ids and counts, never comment or highlighted text",
			}),
		),
		pdf_comment_edits: Type.Optional(
			Type.Array(
				Type.Object(
					{
						id: Type.String({ description: "Comment id from a prior result", minLength: 1 }),
						comment: Type.String({ description: "Replacement comment text", minLength: 1 }),
					},
					{ additionalProperties: false },
				),
				{
					minItems: 1,
					maxItems: 100,
					description:
						"PDF patch: replace the text of 1–100 comments this tool created (ids from pdf_comments results or view=comments); unknown or foreign ids fail the whole call. Can be combined with pdf_comments in one atomic output",
				},
			),
		),
		xlsx_patches: Type.Optional(xlsxPatchesSchema),
		xlsx_formula: Type.Optional({
			...xlsxFormulaSchema,
			description:
				"XLSX patch single formula; mutually exclusive with xlsx_formulas, xlsx_formula_sheets, range, and xlsx_patches. Requires sheet and distinct output; other caches may be stale",
		}),
		xlsx_formulas: Type.Optional(
			Type.Array(xlsxFormulaSchema, {
				minItems: 1,
				maxItems: 1000,
				description:
					"Atomic XLSX formula batch on one sheet, 1–1000 explicit unique cells (case-insensitive). Mutually exclusive with xlsx_formula, xlsx_formula_sheets, range, and xlsx_patches. Requires sheet and distinct output; dry_run validates all targets without writing. Returns cells/count, never formulas or values; other caches may be stale",
			}),
		),
		xlsx_formula_sheets: Type.Optional(
			Type.Array(
				Type.Object(
					{
						sheet: Type.String({ minLength: 1 }),
						formulas: Type.Array(xlsxFormulaSchema, { minItems: 1, maxItems: 1000 }),
					},
					{ additionalProperties: false },
				),
				{
					minItems: 1,
					maxItems: 1000,
					description:
						"Atomic multi-worksheet formula batch: 1–1000 total explicit cells; unique normalized cells per sheet and unique resolved worksheets (names case-insensitive). Mutually exclusive with sheet, xlsx_formula, xlsx_formulas, range, xlsx_patches. Requires distinct output; dry_run validates without writing. Returns sheets with normalized cells/count and total formulaCount, never formulas or values",
				},
			),
		),
		xlsx_recalculate: Type.Optional(
			Type.Boolean({
				description:
					"XLSX recalculation request; must be true. Mutually exclusive with sheet, range, xlsx_patches, xlsx_formula, xlsx_formulas, and xlsx_formula_sheets. Requires distinct output; dry_run preserves output. Recomputes every formula cell's cached value with headless LibreOffice and writes only those cached values; formulas and inputs are unchanged. Returns counts and warnings, never formula text or values. Requires libreoffice on PATH.",
			}),
		),
		dry_run: Type.Optional(
			Type.Boolean({
				default: false,
				description:
					"Validate and report without creating or replacing output; output remains required for DOCX/PPTX patch or XLSX write/patch. Not supported with tracked_replace",
			}),
		),
		overwrite_original: Type.Optional(
			Type.Boolean({
				default: false,
				description:
					"Allow the output to be the input path, replacing the source file on successful publication (DOCX patch with patches or tracked_replace, PPTX patch, XLSX patch, XLSX write, XLSX recalculation). Without it the output must differ from the input. An existing distinct output is always replaced on success; a dry run never writes. Not accepted for DOCX/PPTX Markdown writes.",
			}),
		),
		literal: Type.Optional(
			Type.Boolean({
				default: false,
				description: "Treat a PDF search query as literal text instead of a regular expression",
			}),
		),
		ignore_case: Type.Optional(
			Type.Boolean({ default: false, description: "Ignore case for DOCX or PDF search" }),
		),
		first_page: Type.Optional(
			Type.Integer({
				description:
					"First 1-based inclusive page for PDF read/search or DOCX/PPTX/XLSX image read; must not exceed last_page",
				minimum: 1,
			}),
		),
		last_page: Type.Optional(
			Type.Integer({
				description:
					"Last 1-based inclusive page for PDF read/search or DOCX/PDF/XLSX image read; must not precede first_page",
				minimum: 1,
			}),
		),
	},
	{ additionalProperties: false },
);

export type FillerInput = Static<typeof fillerSchema>;

type ControlField = Exclude<keyof FillerInput, "format" | "action" | "view" | "path">;

const controlFields: ControlField[] = [
	"output",
	"query",
	"track_changes",
	"review_offset",
	"review_limit",
	"reference_docx",
	"pptx_patches",
	"reference_pptx",
	"patches",
	"tracked_replace",
	"source_csv",
	"sheet",
	"start_cell",
	"range",
	"has_header",
	"value_mode",
	"column_map",
	"pdf_comments",
	"pdf_comment_edits",
	"xlsx_patches",
	"xlsx_formula",
	"xlsx_formulas",
	"xlsx_formula_sheets",
	"xlsx_recalculate",
	"dry_run",
	"overwrite_original",
	"literal",
	"ignore_case",
	"first_page",
	"last_page",
];

export function normalizePath(path: string, cwd: string): string {
	const withoutAt = path.startsWith("@") ? path.slice(1) : path;
	if (!withoutAt) throw new Error("Path must not be empty");
	return resolve(cwd, withoutAt);
}

function validatePageRange(input: FillerInput): void {
	if (input.first_page !== undefined && input.first_page < 1) {
		throw new Error("first_page must be at least 1");
	}
	if (input.last_page !== undefined && input.last_page < 1) {
		throw new Error("last_page must be at least 1");
	}
	if (
		input.first_page !== undefined &&
		input.last_page !== undefined &&
		input.first_page > input.last_page
	) {
		throw new Error("first_page cannot be greater than last_page");
	}
}

function assertParameters(input: FillerInput, allowed: readonly ControlField[]): void {
	for (const field of controlFields) {
		if (input[field] !== undefined && !allowed.includes(field)) {
			throw new Error(`${input.format.toUpperCase()} ${input.action} does not accept ${field}`);
		}
	}
}

function assertSupported(input: FillerInput): void {
	if (!input.path) throw new Error("path must not be empty");
	if (input.format === "docx") {
		if (input.action === "read") {
			if (input.view === "review") {
				assertParameters(input, ["review_offset", "review_limit"]);
				return;
			}
			if (input.view === "image") {
				if (!input.output) throw new Error("DOCX image reads require output");
				assertParameters(input, ["output", "first_page", "last_page"]);
				return;
			}
			if (input.view !== "text" && input.view !== "formatting") {
				throw new Error(
					"DOCX read requires view=text, view=formatting, view=review, or view=image",
				);
			}
			assertParameters(input, input.view === "text" ? ["track_changes"] : []);
			return;
		}
		if (input.action === "search") {
			if (input.view !== undefined && input.view !== "text") {
				throw new Error("DOCX search requires view=text or no view");
			}
			if (!input.query) throw new Error("DOCX search requires query");
			assertParameters(input, ["query", "ignore_case", "track_changes"]);
			return;
		}
		if (input.action === "write") {
			if (input.view !== undefined || !input.output) {
				throw new Error("DOCX write requires output and does not accept view");
			}
			assertParameters(input, ["output", "reference_docx"]);
			return;
		}
		if (input.view !== undefined || !input.output) {
			throw new Error("DOCX patch requires output and does not accept view");
		}
		if ((input.patches === undefined) === (input.tracked_replace === undefined)) {
			throw new Error("DOCX patch requires exactly one of patches or tracked_replace");
		}
		if (input.tracked_replace !== undefined) {
			assertParameters(input, ["output", "tracked_replace", "overwrite_original"]);
			validateTrackedReplacement(input.tracked_replace);
		} else {
			assertParameters(input, ["output", "patches", "dry_run", "overwrite_original"]);
		}
		return;
	}

	if (input.format === "pptx") {
		if (input.action === "read") {
			if (input.view === "image") {
				if (!input.output) throw new Error("PPTX image reads require output");
				assertParameters(input, ["output", "first_page", "last_page"]);
				return;
			}
			if (input.view !== "text" && input.view !== "formatting") {
				throw new Error("PPTX read requires view=text, view=formatting, or view=image");
			}
			assertParameters(input, []);
			return;
		}
		if (input.action === "search") {
			if (input.view !== undefined && input.view !== "text") {
				throw new Error("PPTX search requires view=text or no view");
			}
			if (!input.query) throw new Error("PPTX search requires query");
			assertParameters(input, ["query", "ignore_case"]);
			return;
		}
		if (input.action === "write") {
			if (input.view !== undefined || !input.output) {
				throw new Error("PPTX write requires output and does not accept view");
			}
			assertParameters(input, ["output", "reference_pptx"]);
			return;
		}
		if (input.view !== undefined || !input.output) {
			throw new Error("PPTX patch requires output and does not accept view");
		}
		if (input.pptx_patches === undefined) throw new Error("PPTX patch requires pptx_patches");
		assertParameters(input, ["output", "pptx_patches", "dry_run", "overwrite_original"]);
		return;
	}

	if (input.format === "xlsx") {
		if (input.action === "read") {
			if (input.view === "image") {
				if (!input.output) throw new Error("XLSX image reads require output");
				assertParameters(input, ["output", "first_page", "last_page"]);
				return;
			}
			if (input.view !== "structure")
				throw new Error("XLSX read requires view=structure or view=image");
			assertParameters(input, []);
			return;
		}
		if (input.action === "search") throw new Error("XLSX search is not supported");
		if (input.action === "write") {
			if (input.view !== undefined) throw new Error("XLSX write does not accept view");
			if (!input.output || !input.source_csv || !input.sheet || !input.start_cell) {
				throw new Error("XLSX write requires output, source_csv, sheet, and start_cell");
			}
			assertParameters(input, [
				"output",
				"source_csv",
				"sheet",
				"start_cell",
				"has_header",
				"value_mode",
				"column_map",
				"dry_run",
				"overwrite_original",
			]);
			return;
		}
		if (input.view !== undefined) throw new Error("XLSX patch does not accept view");
		if (input.xlsx_recalculate !== undefined) {
			if (!input.output) throw new Error("XLSX recalculation requires output");
			if (input.xlsx_recalculate !== true) throw new Error("xlsx_recalculate must be true");
			assertParameters(input, ["output", "xlsx_recalculate", "dry_run", "overwrite_original"]);
			return;
		}
		if (input.xlsx_formula_sheets !== undefined) {
			if (!input.output) throw new Error("XLSX formula patch requires output");
			assertParameters(input, ["output", "xlsx_formula_sheets", "dry_run", "overwrite_original"]);
			validateXlsxFormulaSheets(input.xlsx_formula_sheets);
			return;
		}
		if (input.xlsx_formula !== undefined || input.xlsx_formulas !== undefined) {
			if (!input.output || !input.sheet)
				throw new Error("XLSX formula patch requires output and sheet");
			if (input.xlsx_formula !== undefined && input.xlsx_formulas !== undefined)
				throw new Error("Use exactly one of xlsx_formula and xlsx_formulas");
			assertParameters(input, [
				"output",
				"sheet",
				"xlsx_formula",
				"xlsx_formulas",
				"dry_run",
				"overwrite_original",
			]);
			if (input.xlsx_formulas !== undefined) validateXlsxFormulas(input.xlsx_formulas);
			else validateXlsxFormula(input.xlsx_formula);
			return;
		}
		if (!input.output || !input.sheet || !input.range || !input.xlsx_patches) {
			throw new Error("XLSX patch requires output, sheet, range, and xlsx_patches");
		}
		assertParameters(input, [
			"output",
			"sheet",
			"range",
			"xlsx_patches",
			"dry_run",
			"overwrite_original",
		]);
		return;
	}

	if (input.action === "write") throw new Error("PDF write operations are not supported");
	if (input.action === "patch") {
		if (input.view !== undefined || !input.output) {
			throw new Error("PDF patch requires output and does not accept view");
		}
		if (input.pdf_comments === undefined && input.pdf_comment_edits === undefined) {
			throw new Error("PDF patch requires pdf_comments or pdf_comment_edits");
		}
		assertParameters(input, ["output", "pdf_comments", "pdf_comment_edits", "dry_run"]);
		return;
	}
	if (input.action === "search") {
		if (input.view !== undefined && input.view !== "text") {
			throw new Error("PDF search requires view=text or no view");
		}
		if (!input.query) throw new Error("PDF search requires query");
		assertParameters(input, ["query", "literal", "ignore_case", "first_page", "last_page"]);
		return;
	}
	if (input.view === "comments") {
		assertParameters(input, []);
		return;
	}
	if (input.view !== "text" && input.view !== "image") {
		throw new Error("PDF read requires view=text, view=image, or view=comments");
	}
	if (input.view === "image") {
		if (!input.output) throw new Error("PDF image reads require output");
		assertParameters(input, ["output", "first_page", "last_page"]);
		return;
	}
	assertParameters(input, ["first_page", "last_page"]);
}

export interface FillerResult {
	text: string;
	details: Record<string, unknown>;
}

export async function executeFiller(
	input: FillerInput,
	cwd: string,
	signal?: AbortSignal,
): Promise<FillerResult> {
	assertSupported(input);
	validatePageRange(input);
	const path = normalizePath(input.path, cwd);
	const range = { firstPage: input.first_page, lastPage: input.last_page };

	if (input.format === "docx") {
		if (input.action === "read" && input.view === "review") {
			const review = await inspectDocxReview(path, input.review_offset, input.review_limit);
			return {
				text: JSON.stringify(review, null, 2),
				details: {
					totalComments: review.totalComments,
					returned: review.returned,
					nextOffset: review.nextOffset,
				},
			};
		}
		if (input.action === "read" && input.view === "formatting") {
			const formatting = await inspectDocx(path);
			return { text: JSON.stringify(formatting, null, 2), details: { formatting } };
		}
		if (input.action === "read" && input.view === "image") {
			const output = normalizePath(input.output ?? "", cwd);
			const temporary = await mkdtemp(join(tmpdir(), "pi-filler-docx-"));
			try {
				const converted = await runCommand(
					"libreoffice",
					[
						`-env:UserInstallation=${pathToFileURL(join(temporary, "profile")).href}`,
						"--headless",
						"--convert-to",
						"pdf",
						"--outdir",
						temporary,
						path,
					],
					{ signal, cwd: temporary },
				);
				if (converted.code !== 0) throw commandError(converted, "libreoffice");
				const pdf = join(temporary, `${basename(path, extname(path))}.pdf`);
				const result = await withFileMutationQueue(output, () =>
					renderPdfPages(pdf, output, range, { signal }),
				);
				return {
					text: `Rendered ${result.length} page image${result.length === 1 ? "" : "s"}:\n${result.join("\n")}`,
					details: { files: result },
				};
			} finally {
				await rm(temporary, { recursive: true, force: true });
			}
		}
		if (input.action === "read") {
			const result = await readDocxText(path, { signal, trackChanges: input.track_changes });
			return { text: result.text, details: { truncation: result.truncation } };
		}
		if (input.action === "search") {
			const result = await searchDocx(path, input.query ?? "", {
				trackChanges: input.track_changes,
				ignoreCase: input.ignore_case,
				signal,
			});
			return {
				text: result.text || "No matches found",
				details: { matchCount: result.matchCount, truncation: result.truncation },
			};
		}
		if (input.action === "patch") {
			const output = normalizePath(input.output ?? "", cwd);
			const result = await withFileMutationQueue(output, async () =>
				input.tracked_replace !== undefined
					? replaceTracked(
							path,
							output,
							input.tracked_replace,
							cwd,
							signal,
							input.overwrite_original === true,
						)
					: patchDocx(
							path,
							output,
							input.patches as DocxPatchSet,
							input.dry_run,
							input.overwrite_original === true,
						),
			);
			return {
				text: JSON.stringify(result, null, 2),
				details: result as unknown as Record<string, unknown>,
			};
		}

		const output = normalizePath(input.output ?? "", cwd);
		if (output === path) throw new Error("DOCX write output must differ from input");
		const referenceDocx = input.reference_docx
			? normalizePath(input.reference_docx, cwd)
			: undefined;
		await withFileMutationQueue(output, () =>
			writeDocx(path, output, { referenceDocx, signal, cwd: dirname(path) }),
		);
		return { text: `Wrote ${output}`, details: { output } };
	}

	if (input.format === "pptx") {
		if (input.action === "read" && input.view === "formatting") {
			const formatting = await inspectPptx(path);
			return { text: JSON.stringify(formatting, null, 2), details: { formatting } };
		}
		if (input.action === "read" && input.view === "image") {
			const output = normalizePath(input.output ?? "", cwd);
			const temporary = await mkdtemp(join(tmpdir(), "pi-filler-pptx-"));
			try {
				const converted = await runCommand(
					"libreoffice",
					[
						`-env:UserInstallation=${pathToFileURL(join(temporary, "profile")).href}`,
						"--headless",
						"--convert-to",
						"pdf",
						"--outdir",
						temporary,
						path,
					],
					{ signal, cwd: temporary },
				);
				if (converted.code !== 0) throw commandError(converted, "libreoffice");
				const pdf = join(temporary, `${basename(path, extname(path))}.pdf`);
				const result = await withFileMutationQueue(output, () =>
					renderPdfPages(pdf, output, range, { signal }),
				);
				return {
					text: `Rendered ${result.length} slide image${result.length === 1 ? "" : "s"}:\n${result.join("\n")}`,
					details: { files: result },
				};
			} finally {
				await rm(temporary, { recursive: true, force: true });
			}
		}
		if (input.action === "read") {
			const result = await readPptxText(path, { signal });
			return { text: result.text, details: { truncation: result.truncation } };
		}
		if (input.action === "search") {
			const result = await searchPptx(path, input.query ?? "", {
				ignoreCase: input.ignore_case,
				signal,
			});
			return {
				text: result.text || "No matches found",
				details: { matchCount: result.matchCount, truncation: result.truncation },
			};
		}
		if (input.action === "patch") {
			const output = normalizePath(input.output ?? "", cwd);
			const result = await withFileMutationQueue(output, () =>
				patchPptx(
					path,
					output,
					input.pptx_patches as PptxPatchSet,
					input.dry_run,
					input.overwrite_original === true,
				),
			);
			return {
				text: JSON.stringify(result, null, 2),
				details: result as unknown as Record<string, unknown>,
			};
		}

		const output = normalizePath(input.output ?? "", cwd);
		if (output === path) throw new Error("PPTX write output must differ from input");
		const referencePptx = input.reference_pptx
			? normalizePath(input.reference_pptx, cwd)
			: undefined;
		await withFileMutationQueue(output, () =>
			writePptx(path, output, { referencePptx, signal, cwd: dirname(path) }),
		);
		return { text: `Wrote ${output}`, details: { output } };
	}

	if (input.format === "xlsx") {
		if (input.action === "read") {
			if (input.view === "image") {
				const output = normalizePath(input.output ?? "", cwd);
				const temporary = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-"));
				try {
					const converted = await runCommand(
						"libreoffice",
						[
							`-env:UserInstallation=${pathToFileURL(join(temporary, "profile")).href}`,
							"--headless",
							"--convert-to",
							"pdf",
							"--outdir",
							temporary,
							path,
						],
						{ signal, cwd: temporary },
					);
					if (converted.code !== 0) throw commandError(converted, "libreoffice");
					const pdf = join(temporary, `${basename(path, extname(path))}.pdf`);
					const files = await withFileMutationQueue(output, () =>
						renderPdfPages(pdf, output, range, { signal }),
					);
					return {
						text: `Rendered ${files.length} page image${files.length === 1 ? "" : "s"}:\n${files.join("\n")}`,
						details: { files },
					};
				} finally {
					await rm(temporary, { recursive: true, force: true });
				}
			}
			const result = await inspectXlsxStructure(path);
			return {
				text: JSON.stringify(result, null, 2),
				details: result as unknown as Record<string, unknown>,
			};
		}
		const output = normalizePath(input.output ?? "", cwd);
		if (input.action === "write") {
			const sourceCsv = normalizePath(input.source_csv ?? "", cwd);
			const result = await withFileMutationQueue(output, () =>
				fillXlsxFromCsv(path, sourceCsv, output, {
					sheet: input.sheet ?? "",
					startCell: input.start_cell ?? "",
					hasHeader: input.has_header,
					valueMode: input.value_mode,
					columnMap: input.column_map,
					dryRun: input.dry_run,
					signal,
					overwriteOriginal: input.overwrite_original === true,
				}),
			);
			return {
				text: JSON.stringify(result, null, 2),
				details: result as unknown as Record<string, unknown>,
			};
		}
		if (input.xlsx_recalculate !== undefined) {
			const result = await withFileMutationQueue(output, () =>
				recalculateXlsx(path, output, input.dry_run, signal, input.overwrite_original === true),
			);
			return { text: JSON.stringify(result, null, 2), details: { ...result } };
		}
		if (input.xlsx_formula_sheets !== undefined) {
			const result = await withFileMutationQueue(output, () =>
				patchXlsxFormulaSheets(
					path,
					output,
					input.xlsx_formula_sheets ?? [],
					input.dry_run,
					signal,
					input.overwrite_original === true,
				),
			);
			return { text: JSON.stringify(result, null, 2), details: { ...result } };
		}
		const formula = input.xlsx_formulas ?? input.xlsx_formula;
		if (formula !== undefined) {
			const result = await withFileMutationQueue(output, () =>
				patchXlsxFormula(
					path,
					output,
					input.sheet ?? "",
					formula,
					input.dry_run,
					signal,
					input.overwrite_original === true,
				),
			);
			return { text: JSON.stringify(result, null, 2), details: { ...result } };
		}
		const result = await withFileMutationQueue(output, () =>
			patchXlsx(
				path,
				output,
				{
					sheet: input.sheet ?? "",
					range: input.range ?? "",
					patches: input.xlsx_patches as XlsxPatchSet,
				},
				input.dry_run,
				signal,
				input.overwrite_original === true,
			),
		);
		return {
			text: JSON.stringify(result, null, 2),
			details: result as unknown as Record<string, unknown>,
		};
	}

	if (input.action === "patch") {
		const output = normalizePath(input.output ?? "", cwd);
		if (output === path) throw new Error("PDF patch output must differ from input");
		const result = await withFileMutationQueue(output, () =>
			annotatePdf(
				path,
				output,
				{
					comments: input.pdf_comments as PdfComment[] | undefined,
					edits: input.pdf_comment_edits,
				},
				input.dry_run === true,
				{ signal },
			),
		);
		return {
			text: JSON.stringify(result, null, 2),
			details: result as unknown as Record<string, unknown>,
		};
	}
	if (input.action === "search") {
		const result = await searchPdf(path, input.query ?? "", {
			...range,
			ignoreCase: input.ignore_case,
			literal: input.literal,
			signal,
		});
		const truncation = truncateHead(result.text);
		return {
			text: truncation.truncated
				? `${truncation.content}\n\n[Search output truncated at ${DEFAULT_MAX_BYTES} bytes]`
				: truncation.content || "No matches found",
			details: { matchCount: result.matchCount, truncation },
		};
	}
	if (input.view === "comments") {
		const comments = await listPdfComments(path);
		return {
			text: JSON.stringify(comments, null, 2),
			details: { commentCount: comments.length },
		};
	}
	if (input.view === "text") {
		const result = await readPdfText(path, range, { signal });
		return { text: result.text, details: { truncation: result.truncation } };
	}

	const output = normalizePath(input.output ?? "", cwd);
	const result = await withFileMutationQueue(output, () =>
		renderPdfPages(path, output, range, { signal }),
	);
	return {
		text: `Rendered ${result.length} page image${result.length === 1 ? "" : "s"}:\n${result.join("\n")}`,
		details: { files: result },
	};
}

export default function extension(pi: ExtensionAPI): void {
	pi.registerTool({
		name: "filler",
		label: "Filler",
		description:
			"Work directly with local DOCX, PDF, PPTX, and XLSX files. Read, search, or render DOCX/PPTX and PDF; write DOCX/PPTX from Markdown; add PDF highlight comments to existing text (pdf_comments); inspect or render XLSX, fill XLSX templates from CSV, patch supported formatting, set ordinary XLSX formulas atomically without evaluation, or recalculate cached XLSX formula values with LibreOffice (xlsx_recalculate).",
		promptSnippet:
			"Work directly with local DOCX, PDF, PPTX, and XLSX files using supported format/action/view combinations.",
		promptGuidelines: [
			"Use filler when a supported deterministic local DOCX, PDF, PPTX, or XLSX operation satisfies the task; choose a supported format/action/view combination and provide only that operation's fields.",
			"For filler reads, use DOCX text, formatting, review, or image; PPTX text, formatting, or image; PDF text or image; or XLSX structure or image. DOCX review returns bounded comment text, authors, main-document anchor excerpts, and supported legacy reply/resolution metadata; review_offset defaults to 0 and review_limit to 20 (max 50). Follow nextOffset; text/excerpts may be truncated and modern thread metadata may be unsupported.",
			"For filler searches, use DOCX, PPTX, or PDF text (or omit view); XLSX search is unsupported.",
			"For filler PDF comments, use format pdf, action patch, path (input), output (new distinct file), and pdf_comments [{page, text, comment, author}, ...] (1–100 items, 1-based page). text is the exact case-sensitive phrase to highlight and must match exactly once on that page; comment and author are stored in the highlight annotation. Results return new comment ids. To revise comments this tool created, use pdf_comment_edits [{id, comment}, ...] (ids from results or view=comments); it can be combined with pdf_comments in one call. Use format pdf, action read, view comments to list only this tool's comments with ids, pages, author, comment, anchor, and modified date. Only unrotated, unsigned, unencrypted PDFs are supported; comments are not deletable and replies are unsupported. Results return counts, pages, and ids, but never highlighted or comment text in patch results. dry_run validates without writing. PDF write is unsupported.",
			"Existing distinct outputs are replaced on success. Set overwrite_original: true only when the operator or agent intends to replace the source file itself (DOCX patch, PPTX patch, XLSX patch, XLSX write, XLSX recalculation); a dry run never writes. Without it the output must differ from the input.",
			"For filler image output, use an existing directory or PNG prefix; the operation creates numbered page/slide images, not one final PNG. DOCX, PPTX, and XLSX images convert through LibreOffice; XLSX images render workbook print pages and do not accept sheet. first_page and last_page select resulting pages.",
			"For DOCX/PPTX writes, path is Markdown and output is the new document; reference_docx/reference_pptx is optional. DOCX uses Pandoc Markdown (smart typography disabled), including native review spans; PPTX uses GFM Markdown. DOCX generation is not a lossless round trip of an existing Word layout. For XLSX writes, path is the template and source_csv, sheet, A1-style start_cell, and output are required; column_map optionally maps all CSV ordinals or all exact headers to XLSX columns.",
			"When an XLSX write or patch can perform the requested deterministic transformation without inspecting CSV or cell contents, do not read those contents into model context merely to carry out the transformation. Processing stays local and results do not return cell or CSV contents.",
			"For filler patches, provide a new output path plus DOCX patches, PPTX pptx_patches, or XLSX sheet, A1-style range, and xlsx_patches. PPTX text replacements target slide text paragraphs and may name a 1-based slide. dry_run: true still requires output but does not create or replace it.",
			"Use filler XLSX patch with output and xlsx_recalculate: true alone to recompute all formula cached values with headless LibreOffice (requires libreoffice on PATH); it writes only cached values, never formulas or results, and refuses external links. Its results are LibreOffice results, not Excel results.",
			"Use filler XLSX patch with sheet and either xlsx_formula {cell, formula} or xlsx_formulas [{cell, formula}, ...], mutually exclusive with each other and range/xlsx_patches. Alternatively use xlsx_formula_sheets [{sheet, formulas:[{cell, formula}, ...]}, ...] without sheet/xlsx_formula/xlsx_formulas/range/xlsx_patches for one atomic multi-worksheet edit: 1–1000 total targets, unique case-insensitive resolved worksheets; returns sheets with normalized cells and per-sheet formulaCount plus total formulaCount. Batches contain 1–1000 explicit cells on one sheet, with no duplicate case-normalized addresses; every target is validated before atomic publication. Creates absent cells or replaces existing values/formulas, preserving styles. cell must be one explicit unmerged A1 address; formula must start with =, have a nonblank body, use valid XML characters without CR, and be at most 8192 UTF-16 code units. Batch results contain normalized cells and formulaCount, never values/formulas. Requires distinct output; dry_run preserves output. Stores an ordinary formula without evaluating it, requests full automatic recalculation on open, and leaves other formula caches unchanged (possibly stale). Conventional xl/calcChain.xml calculation order caches are discarded with their relationship/content type (included in changedParts); nonstandard or ambiguous chains are rejected. This does not calculate dependent caches. Signed packages, shared/array/dataTable formulas, ambiguous rows/cells, and unsupported formula metadata are rejected. Formula and workbook values stay local and are not returned.",
			"DOCX patch alternatively accepts exactly one tracked_replace {old, new, author, expected_sha256?}, not patches. Requires independently installed docxquisite on PATH and a distinct output (an existing output is replaced on success), no view or dry_run (even false). old and author must be nonempty; empty new deletes. The optional source SHA-256 is lowercase 64 hex. Restricted to one unique single-run main-story match; existing revisions and complex documents may be rejected. Processing stays local; requested text/author and toolkit diagnostics are not returned. Structural checks do not prove Word rendering. Transport/protocol failure may leave an output: inspect it before retrying; no automatic cleanup or retry.",
			"For filler DOCX text reads/searches, track_changes defaults to accept; use reject for the pre-edit text or all for insertion/deletion spans and comments. This selects an extraction view, not native Word acceptance/rejection. For filler defaults, PDF queries are regular expressions by default; use literal: true for literal text. XLSX writes skip a header and preserve text by default.",
			"For filler results, DOCX/PPTX search may be truncated; narrow its query. PDF text or search may be truncated; narrow its page range or query.",
		],
		parameters: fillerSchema,
		async execute(_toolCallId, input, signal, _onUpdate, ctx: ExtensionContext) {
			const result = await executeFiller(input, ctx.cwd, signal);
			return { content: [{ type: "text", text: result.text }], details: result.details };
		},
	});
}
