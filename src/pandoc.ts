import { randomUUID } from "node:crypto";
import { mkdir, rename, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { type TruncationResult, truncateHead } from "@earendil-works/pi-coding-agent";
import { inspectDocx } from "./docx.js";
import { inspectPptx } from "./pptx.js";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";

export interface DocxTextResult {
	text: string;
	truncation: TruncationResult;
}

export interface DocxSearchResult {
	matchCount: number;
	text: string;
	truncation: TruncationResult;
}

function commandOptions(options: RunCommandOptions | undefined, path: string): RunCommandOptions {
	return { ...options, cwd: options?.cwd ?? dirname(path) };
}

async function runPandoc(
	args: string[],
	options: RunCommandOptions | undefined,
	path: string,
): Promise<Buffer> {
	const result = await runCommand("pandoc", args, commandOptions(options, path));
	if (result.code !== 0) throw commandError(result, "pandoc");
	return result.stdout;
}

async function extractMarkdown(
	path: string,
	from: "docx" | "pptx",
	options?: RunCommandOptions,
): Promise<string> {
	const source = resolve(path);
	const output = await runPandoc(
		[`--from=${from}`, "--to=gfm", "--wrap=none", source],
		options,
		source,
	);
	return output.toString("utf8");
}

async function extractDocxMarkdown(path: string, options?: RunCommandOptions): Promise<string> {
	return extractMarkdown(path, "docx", options);
}

async function extractPptxMarkdown(path: string, options?: RunCommandOptions): Promise<string> {
	return extractMarkdown(path, "pptx", options);
}

export async function readDocxText(
	path: string,
	options?: RunCommandOptions,
): Promise<DocxTextResult> {
	const truncation = truncateHead(await extractDocxMarkdown(path, options));
	return { text: truncation.content, truncation };
}

function searchMarkdown(extracted: string, query: string, ignoreCase = false): DocxSearchResult {
	const needle = ignoreCase ? query.toLocaleLowerCase() : query;
	const matches: string[] = [];
	for (const [index, text] of extracted.split(/\r?\n/).entries()) {
		const haystack = ignoreCase ? text.toLocaleLowerCase() : text;
		if (haystack.includes(needle)) matches.push(`line ${index + 1}: ${text}`);
	}
	const truncation = truncateHead(matches.join("\n"));
	return { matchCount: matches.length, text: truncation.content, truncation };
}

export async function searchDocx(
	path: string,
	query: string,
	options: RunCommandOptions & { ignoreCase?: boolean } = {},
): Promise<DocxSearchResult> {
	return searchMarkdown(await extractDocxMarkdown(path, options), query, options.ignoreCase);
}

export async function readPptxText(
	path: string,
	options?: RunCommandOptions,
): Promise<DocxTextResult> {
	const truncation = truncateHead(await extractPptxMarkdown(path, options));
	return { text: truncation.content, truncation };
}

export async function searchPptx(
	path: string,
	query: string,
	options: RunCommandOptions & { ignoreCase?: boolean } = {},
): Promise<DocxSearchResult> {
	return searchMarkdown(await extractPptxMarkdown(path, options), query, options.ignoreCase);
}

export interface WriteDocxOptions extends RunCommandOptions {
	referenceDocx?: string;
}

export interface WritePptxOptions extends RunCommandOptions {
	referencePptx?: string;
}

export async function writeDocx(
	sourcePath: string,
	outputPath: string,
	options: WriteDocxOptions = {},
): Promise<void> {
	const source = resolve(sourcePath);
	const output = resolve(outputPath);
	if (source === output) throw new Error("DOCX write output must differ from input");
	const temporary = join(
		dirname(output),
		`.${output.split(/[\\/]/).pop() ?? "output.docx"}.${randomUUID()}.tmp`,
	);
	await mkdir(dirname(output), { recursive: true });
	const args = ["--from=gfm", "--to=docx", "--output", temporary];
	if (options.referenceDocx) args.push("--reference-doc", resolve(options.referenceDocx));
	args.push(source);
	try {
		const result = await runCommand("pandoc", args, {
			...options,
			cwd: options.cwd ?? dirname(source),
		});
		if (result.code !== 0) throw commandError(result, "pandoc");
		await inspectDocx(temporary);
		await rename(temporary, output);
	} finally {
		await rm(temporary, { force: true });
	}
}

export async function writePptx(
	sourcePath: string,
	outputPath: string,
	options: WritePptxOptions = {},
): Promise<void> {
	const source = resolve(sourcePath);
	const output = resolve(outputPath);
	if (source === output) throw new Error("PPTX write output must differ from input");
	const temporary = join(
		dirname(output),
		`.${output.split(/[\\/]/).pop() ?? "output.pptx"}.${randomUUID()}.tmp`,
	);
	await mkdir(dirname(output), { recursive: true });
	const args = ["--from=gfm", "--to=pptx", "--output", temporary];
	if (options.referencePptx) args.push("--reference-doc", resolve(options.referencePptx));
	args.push(source);
	try {
		const result = await runCommand("pandoc", args, {
			...options,
			cwd: options.cwd ?? dirname(source),
		});
		if (result.code !== 0) throw commandError(result, "pandoc");
		await inspectPptx(temporary);
		await rename(temporary, output);
	} finally {
		await rm(temporary, { force: true });
	}
}
