import { randomUUID } from "node:crypto";
import { mkdir, rename, rm } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { type TruncationResult, truncateHead } from "@earendil-works/pi-coding-agent";
import { inspectDocx } from "./docx.js";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";

export interface DocxTextResult {
	text: string;
	truncation: TruncationResult;
}

export interface DocxSearchMatch {
	line: number;
	text: string;
}

export interface DocxSearchResult {
	matches: DocxSearchMatch[];
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

async function extractDocxMarkdown(path: string, options?: RunCommandOptions): Promise<string> {
	const output = await runPandoc(["--from=docx", "--to=gfm", "--wrap=none", path], options, path);
	return output.toString("utf8");
}

export async function readDocxText(
	path: string,
	options?: RunCommandOptions,
): Promise<DocxTextResult> {
	const truncation = truncateHead(await extractDocxMarkdown(path, options));
	return { text: truncation.content, truncation };
}

export async function searchDocx(
	path: string,
	query: string,
	options: RunCommandOptions & { ignoreCase?: boolean } = {},
): Promise<DocxSearchResult> {
	const extracted = await extractDocxMarkdown(path, options);
	const needle = options.ignoreCase ? query.toLocaleLowerCase() : query;
	const matches = extracted
		.split(/\r?\n/)
		.map((text, index) => ({ line: index + 1, text }))
		.filter(({ text }) => {
			const haystack = options.ignoreCase ? text.toLocaleLowerCase() : text;
			return haystack.includes(needle);
		});
	const result = matches.map((match) => `line ${match.line}: ${match.text}`).join("\n");
	const truncation = truncateHead(result);
	return { matches, text: truncation.content, truncation };
}

export interface WriteDocxOptions extends RunCommandOptions {
	referenceDocx?: string;
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
	if (options.referenceDocx) args.push("--reference-doc", options.referenceDocx);
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
