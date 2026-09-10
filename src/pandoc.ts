import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm, stat } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { type TruncationResult, truncateHead } from "@earendil-works/pi-coding-agent";
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

export async function readDocxText(
	path: string,
	options?: RunCommandOptions,
): Promise<DocxTextResult> {
	const output = await runPandoc(["--from=docx", "--to=gfm", "--wrap=none", path], options, path);
	const truncation = truncateHead(output.toString("utf8"));
	return { text: truncation.content, truncation };
}

export async function searchDocx(
	path: string,
	query: string,
	options: RunCommandOptions & { ignoreCase?: boolean } = {},
): Promise<DocxSearchResult> {
	const extracted = await readDocxText(path, options);
	const needle = options.ignoreCase ? query.toLocaleLowerCase() : query;
	const matches = extracted.text
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
	const markdown = await readFile(sourcePath, "utf8");
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
	args.push("-");
	try {
		const result = await runCommand("pandoc", args, {
			...options,
			cwd: options.cwd ?? dirname(sourcePath),
			input: markdown,
		});
		if (result.code !== 0) throw commandError(result, "pandoc");
		await stat(temporary);
		await rename(temporary, output);
	} finally {
		await rm(temporary, { force: true });
	}
}
