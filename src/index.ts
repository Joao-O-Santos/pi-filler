import { dirname, resolve } from "node:path";
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

export const fillerSchema = Type.Object({
	format: StringEnum(["docx", "pdf"] as const),
	action: StringEnum(["read", "search", "write", "patch"] as const),
	view: Type.Optional(StringEnum(["text", "formatting", "image"] as const)),
	path: Type.String({ description: "Input document path" }),
	output: Type.Optional(Type.String({ description: "Output file or image prefix" })),
	query: Type.Optional(Type.String({ description: "Search query" })),
	reference_docx: Type.Optional(Type.String({ description: "Reference DOCX for writes" })),
	patches: Type.Optional(Type.Unknown({ description: "Strict typed DOCX formatting patches" })),
	dry_run: Type.Optional(Type.Boolean({ description: "Validate and report without writing" })),
	literal: Type.Optional(Type.Boolean({ description: "Treat a search query as literal text" })),
	ignore_case: Type.Optional(Type.Boolean({ description: "Ignore case while searching" })),
	first_page: Type.Optional(Type.Integer({ minimum: 1 })),
	last_page: Type.Optional(Type.Integer({ minimum: 1 })),
});

export type FillerInput = Static<typeof fillerSchema>;

export function normalizePath(path: string, cwd: string): string {
	const withoutAt = path.startsWith("@") ? path.slice(1) : path;
	return resolve(cwd, withoutAt || ".");
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

function validatePatches(value: unknown): DocxPatchSet {
	if (!value || typeof value !== "object" || Array.isArray(value)) {
		throw new Error("DOCX patches must be an object");
	}
	const input = value as Record<string, unknown>;
	const checkObject = (
		name: string,
		fields: Record<string, string>,
	): Record<string, unknown> | undefined => {
		const candidate = input[name];
		if (candidate === undefined) return undefined;
		if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
			throw new Error(`DOCX ${name} patch must be an object`);
		}
		const object = candidate as Record<string, unknown>;
		for (const key of Object.keys(object)) {
			if (!(key in fields)) throw new Error(`Unknown DOCX patch: ${name}.${key}`);
			if (typeof object[key] !== fields[key]) {
				throw new Error(`DOCX ${name}.${key} must be ${fields[key]}`);
			}
		}
		return object;
	};
	const allowed = new Set([
		"page",
		"margins",
		"lineNumbering",
		"pageNumbering",
		"metadata",
		"anonymize",
	]);
	for (const key of Object.keys(input))
		if (!allowed.has(key)) throw new Error(`Unknown DOCX patch: ${key}`);
	if (input.anonymize !== undefined && typeof input.anonymize !== "boolean") {
		throw new Error("DOCX anonymize must be boolean");
	}
	const page = checkObject("page", { width: "number", height: "number", orientation: "string" });
	if (
		page?.orientation !== undefined &&
		!["portrait", "landscape"].includes(page.orientation as string)
	) {
		throw new Error("DOCX page.orientation must be portrait or landscape");
	}
	checkObject("margins", {
		top: "number",
		bottom: "number",
		left: "number",
		right: "number",
		header: "number",
		footer: "number",
		gutter: "number",
	});
	const line = checkObject("lineNumbering", {
		mode: "string",
		start: "number",
		count_by: "number",
	});
	if (
		line?.mode !== undefined &&
		!["off", "continuous", "restart", "newPage"].includes(line.mode as string)
	) {
		throw new Error("DOCX lineNumbering.mode is invalid");
	}
	checkObject("pageNumbering", { start: "number", format: "string" });
	checkObject("metadata", {
		title: "string",
		subject: "string",
		creator: "string",
		keywords: "string",
		description: "string",
		lastModifiedBy: "string",
	});
	return value as DocxPatchSet;
}

function assertSupported(input: FillerInput): void {
	if (input.format === "docx") {
		if (input.action === "read" && input.view !== "text" && input.view !== "formatting") {
			throw new Error("DOCX read requires view=text or view=formatting");
		}
		if (input.action === "search" && input.view !== undefined && input.view !== "text") {
			throw new Error("DOCX search requires view=text or no view");
		}
		if (input.action === "write" && (input.view !== undefined || !input.output)) {
			throw new Error("DOCX write requires output and does not accept view");
		}
		if (input.action === "search" && !input.query) throw new Error("DOCX search requires query");
		if (input.action === "patch" && (input.view !== undefined || !input.output)) {
			throw new Error("DOCX patch requires output and does not accept view");
		}
		if (input.action === "patch" && input.patches === undefined) {
			throw new Error("DOCX patch requires patches");
		}
		return;
	}
	if (input.action === "read" && input.view !== "text" && input.view !== "image") {
		throw new Error("PDF read requires view=text or view=image");
	}
	if (input.action === "search" && input.view === "image") {
		throw new Error("PDF search does not support view=image");
	}
	if (input.action === "write" || input.action === "patch") {
		throw new Error("PDF write and patch operations are not supported");
	}
	if (input.action === "search" && !input.query) throw new Error("PDF search requires query");
	if (input.action === "read" && input.view === "image" && !input.output) {
		throw new Error("PDF image reads require output");
	}
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

	if (input.format === "docx") {
		if (input.action === "read" && input.view === "formatting") {
			const formatting = await inspectDocx(path);
			return { text: JSON.stringify(formatting, null, 2), details: { formatting } };
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
				details: { matches: result.matches, truncation: result.truncation },
			};
		}
		if (input.action === "patch") {
			const output = normalizePath(input.output ?? "", cwd);
			const result = await withFileMutationQueue(output, () =>
				patchDocx(path, output, validatePatches(input.patches), input.dry_run),
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

	const range = { firstPage: input.first_page, lastPage: input.last_page };
	if (input.action === "search") {
		const result = await searchPdf(path, input.query ?? "", {
			ignoreCase: input.ignore_case,
			literal: input.literal,
			signal,
		});
		const truncation = truncateHead(result.text);
		return {
			text: truncation.truncated
				? `${truncation.content}\n\n[Search output truncated at ${DEFAULT_MAX_BYTES} bytes]`
				: truncation.content || "No matches found",
			details: { matches: result.matches, truncation },
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
		description: "Read, search, inspect, write, or patch DOCX/PDF documents with bounded output.",
		promptSnippet: "Read, search, inspect, write, or patch a document",
		parameters: fillerSchema,
		async execute(_toolCallId, input, signal, _onUpdate, ctx: ExtensionContext) {
			const result = await executeFiller(input, ctx.cwd, signal);
			return { content: [{ type: "text", text: result.text }], details: result.details };
		},
	});
}
