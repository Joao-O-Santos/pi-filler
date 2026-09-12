import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, extname, join, resolve } from "node:path";
import { StringEnum } from "@earendil-works/pi-ai";
import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import {
	DEFAULT_MAX_BYTES,
	truncateHead,
	withFileMutationQueue,
} from "@earendil-works/pi-coding-agent";
import { type Static, Type } from "typebox";
import { type DocxPatchSet, inspectDocx, patchDocx } from "./docx.js";
import { readDocxText, searchDocx, writeDocx } from "./pandoc.js";
import { readPdfText, renderPdfPages, searchPdf } from "./pdf.js";
import { commandError, runCommand } from "./process.js";
import { fillXlsxFromCsv, inspectXlsxStructure, patchXlsx, type XlsxPatchSet } from "./xlsx.js";

const docxPatchesSchema = Type.Object(
	{
		page: Type.Optional(
			Type.Object(
				{
					width: Type.Optional(Type.Integer({ minimum: 1 })),
					height: Type.Optional(Type.Integer({ minimum: 1 })),
					orientation: Type.Optional(StringEnum(["portrait", "landscape"] as const)),
				},
				{ additionalProperties: false, minProperties: 1 },
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
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		lineNumbering: Type.Optional(
			Type.Object(
				{
					mode: Type.Optional(StringEnum(["off", "continuous", "newPage", "newSection"] as const)),
					start: Type.Optional(Type.Integer({ minimum: 0 })),
					count_by: Type.Optional(Type.Integer({ minimum: 1 })),
				},
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		pageNumbering: Type.Optional(
			Type.Object(
				{
					start: Type.Optional(Type.Integer({ minimum: 0 })),
					format: Type.Optional(Type.String({ minLength: 1 })),
				},
				{ additionalProperties: false, minProperties: 1 },
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
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		clearCoreMetadata: Type.Optional(
			Type.Boolean({
				description: "Clear common fields in docProps/core.xml; not comment or revision authors",
			}),
		),
	},
	{ additionalProperties: false, minProperties: 1 },
);

const xlsxPatchesSchema = Type.Object(
	{
		font: Type.Optional(
			Type.Object(
				{
					bold: Type.Optional(Type.Boolean()),
					italic: Type.Optional(Type.Boolean()),
					size: Type.Optional(Type.Number({ exclusiveMinimum: 0 })),
				},
				{ additionalProperties: false, minProperties: 1 },
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
					horizontal: Type.Optional(StringEnum(["left", "center", "right"] as const)),
					vertical: Type.Optional(StringEnum(["top", "center", "bottom"] as const)),
					wrap_text: Type.Optional(Type.Boolean()),
				},
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		number_format: Type.Optional(Type.String({ minLength: 1 })),
		border: Type.Optional(
			Type.Object(
				{ style: Type.Optional(StringEnum(["thin", "medium", "thick"] as const)) },
				{ additionalProperties: false, minProperties: 1 },
			),
		),
		column_width: Type.Optional(Type.Number({ exclusiveMinimum: 0, maximum: 255 })),
		row_height: Type.Optional(Type.Number({ exclusiveMinimum: 0, maximum: 409 })),
	},
	{ additionalProperties: false, minProperties: 1 },
);

export const fillerSchema = Type.Object(
	{
		format: StringEnum(["docx", "pdf", "xlsx"] as const),
		action: StringEnum(["read", "search", "write", "patch"] as const),
		view: Type.Optional(StringEnum(["text", "formatting", "image", "structure"] as const)),
		path: Type.String({ description: "Input document path", minLength: 1 }),
		output: Type.Optional(
			Type.String({ description: "Output file or image prefix", minLength: 1 }),
		),
		query: Type.Optional(Type.String({ description: "Search query", minLength: 1 })),
		reference_docx: Type.Optional(
			Type.String({ description: "Reference DOCX for writes", minLength: 1 }),
		),
		patches: Type.Optional(docxPatchesSchema),
		source_csv: Type.Optional(Type.String({ description: "Local CSV source path", minLength: 1 })),
		sheet: Type.Optional(Type.String({ description: "Worksheet name", minLength: 1 })),
		start_cell: Type.Optional(
			Type.String({ description: "Upper-left XLSX destination cell", minLength: 1 }),
		),
		range: Type.Optional(Type.String({ description: "XLSX target cell range", minLength: 1 })),
		has_header: Type.Optional(Type.Boolean()),
		value_mode: Type.Optional(StringEnum(["text", "auto"] as const)),
		xlsx_patches: Type.Optional(xlsxPatchesSchema),
		dry_run: Type.Optional(Type.Boolean({ description: "Validate and report without writing" })),
		literal: Type.Optional(
			Type.Boolean({ description: "Treat a PDF search query as literal text" }),
		),
		ignore_case: Type.Optional(Type.Boolean({ description: "Ignore case while searching" })),
		first_page: Type.Optional(Type.Integer({ minimum: 1 })),
		last_page: Type.Optional(Type.Integer({ minimum: 1 })),
	},
	{ additionalProperties: false },
);

export type FillerInput = Static<typeof fillerSchema>;

type ControlField = Exclude<keyof FillerInput, "format" | "action" | "view" | "path">;

const controlFields: ControlField[] = [
	"output",
	"query",
	"reference_docx",
	"patches",
	"source_csv",
	"sheet",
	"start_cell",
	"range",
	"has_header",
	"value_mode",
	"xlsx_patches",
	"dry_run",
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
			if (input.view === "image") {
				if (!input.output) throw new Error("DOCX image reads require output");
				assertParameters(input, ["output", "first_page", "last_page"]);
				return;
			}
			if (input.view !== "text" && input.view !== "formatting") {
				throw new Error("DOCX read requires view=text, view=formatting, or view=image");
			}
			assertParameters(input, []);
			return;
		}
		if (input.action === "search") {
			if (input.view !== undefined && input.view !== "text") {
				throw new Error("DOCX search requires view=text or no view");
			}
			if (!input.query) throw new Error("DOCX search requires query");
			assertParameters(input, ["query", "ignore_case"]);
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
		if (input.patches === undefined) throw new Error("DOCX patch requires patches");
		assertParameters(input, ["output", "patches", "dry_run"]);
		return;
	}

	if (input.format === "xlsx") {
		if (input.action === "read") {
			if (input.view !== "structure") throw new Error("XLSX read requires view=structure");
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
				"dry_run",
			]);
			return;
		}
		if (input.view !== undefined) throw new Error("XLSX patch does not accept view");
		if (!input.output || !input.sheet || !input.range || !input.xlsx_patches) {
			throw new Error("XLSX patch requires output, sheet, range, and xlsx_patches");
		}
		assertParameters(input, ["output", "sheet", "range", "xlsx_patches", "dry_run"]);
		return;
	}

	if (input.action === "write" || input.action === "patch") {
		throw new Error("PDF write and patch operations are not supported");
	}
	if (input.action === "search") {
		if (input.view !== undefined && input.view !== "text") {
			throw new Error("PDF search requires view=text or no view");
		}
		if (!input.query) throw new Error("PDF search requires query");
		assertParameters(input, ["query", "literal", "ignore_case", "first_page", "last_page"]);
		return;
	}
	if (input.view !== "text" && input.view !== "image") {
		throw new Error("PDF read requires view=text or view=image");
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
					["--headless", "--convert-to", "pdf", "--outdir", temporary, path],
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
			const result = await readDocxText(path, { signal });
			return { text: result.text, details: { truncation: result.truncation } };
		}
		if (input.action === "search") {
			const result = await searchDocx(path, input.query ?? "", {
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
				patchDocx(path, output, input.patches as DocxPatchSet, input.dry_run),
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

	if (input.format === "xlsx") {
		if (input.action === "read") {
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
					dryRun: input.dry_run,
					signal,
				}),
			);
			return {
				text: JSON.stringify(result, null, 2),
				details: result as unknown as Record<string, unknown>,
			};
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
			"Read, inspect, write, or patch DOCX/PDF/XLSX files. For XLSX, use structural inspection, CSV-to-template filling, and formatting. Spreadsheet contents are processed locally and are not returned to you.",
		promptSnippet:
			"For spreadsheets: read/structure inspects shape, write fills a template from CSV, and patch changes formatting or layout without exposing contents",
		parameters: fillerSchema,
		async execute(_toolCallId, input, signal, _onUpdate, ctx: ExtensionContext) {
			const result = await executeFiller(input, ctx.cwd, signal);
			return { content: [{ type: "text", text: result.text }], details: result.details };
		},
	});
}
