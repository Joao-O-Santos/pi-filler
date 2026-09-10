import { resolve } from "node:path";
import { StringEnum } from "@earendil-works/pi-ai";
import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import {
	DEFAULT_MAX_BYTES,
	truncateHead,
	withFileMutationQueue,
} from "@earendil-works/pi-coding-agent";
import { type Static, Type } from "typebox";
import { readPdfText, renderPdfPages, searchPdf } from "./pdf.js";

export const fillerSchema = Type.Object({
	format: StringEnum(["docx", "pdf"] as const),
	action: StringEnum(["read", "search", "write", "patch"] as const),
	view: Type.Optional(StringEnum(["text", "formatting", "image"] as const)),
	path: Type.String({ description: "Input document path" }),
	output: Type.Optional(Type.String({ description: "Output file or image prefix" })),
	query: Type.Optional(Type.String({ description: "Search query" })),
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

function assertSupported(input: FillerInput): void {
	if (input.format === "docx") {
		throw new Error(
			"DOCX operations are not available yet; supported operations are PDF read, search, and image rendering.",
		);
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
	if (input.action === "search" && !input.query) {
		throw new Error("PDF search requires query");
	}
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
		description:
			"Read and search PDFs or render selected PDF pages. Output is bounded to 50KB and 2000 lines.",
		promptSnippet: "Read, search, or render a PDF document",
		parameters: fillerSchema,
		async execute(_toolCallId, input, signal, _onUpdate, ctx: ExtensionContext) {
			const result = await executeFiller(input, ctx.cwd, signal);
			return { content: [{ type: "text", text: result.text }], details: result.details };
		},
	});
}
