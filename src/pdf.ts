import { mkdir, mkdtemp, readdir, rename, rm, stat } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";
import { type TruncationResult, truncateHead } from "@earendil-works/pi-coding-agent";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";

export interface PageRange {
	firstPage?: number;
	lastPage?: number;
}

export interface PdfTextResult {
	text: string;
	truncation: TruncationResult;
}

export interface PdfSearchMatch {
	page: number;
	text: string;
}

export interface PdfSearchResult {
	matches: PdfSearchMatch[];
	text: string;
}

const PDFGREP_MAX_PAGE = 2_147_483_647;

function pageArgs(range: PageRange): string[] {
	const args: string[] = [];
	if (range.firstPage !== undefined) args.push("-f", String(range.firstPage));
	if (range.lastPage !== undefined) args.push("-l", String(range.lastPage));
	return args;
}

function searchPageArgs(range: PageRange): string[] {
	if (range.firstPage === undefined && range.lastPage === undefined) return [];
	const first = range.firstPage ?? 1;
	const last = range.lastPage ?? PDFGREP_MAX_PAGE;
	return ["--page-range", first === last ? String(first) : `${first}-${last}`];
}

function commandOptions(options: RunCommandOptions | undefined, cwd: string): RunCommandOptions {
	return { ...options, cwd: options?.cwd ?? cwd };
}

export async function readPdfText(
	path: string,
	range: PageRange = {},
	options?: RunCommandOptions,
): Promise<PdfTextResult> {
	const source = resolve(path);
	const result = await runCommand(
		"pdftotext",
		["-layout", ...pageArgs(range), source, "-"],
		commandOptions(options, dirname(source)),
	);
	if (result.code !== 0) throw commandError(result, "pdftotext");
	const truncation = truncateHead(result.stdout.toString("utf8"));
	return { text: truncation.content, truncation };
}

function escapeRegex(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function parseSearchOutput(output: string): PdfSearchMatch[] {
	return output
		.split(/\r?\n/)
		.filter(Boolean)
		.map((line) => {
			const match = line.match(/:(\d+):(.*)$/);
			if (!match) return { page: 0, text: line };
			return { page: Number(match[1]), text: match[2] ?? "" };
		});
}

export async function searchPdf(
	path: string,
	query: string,
	options: RunCommandOptions & PageRange & { ignoreCase?: boolean; literal?: boolean } = {},
): Promise<PdfSearchResult> {
	const source = resolve(path);
	const args = ["--color", "never", "--with-filename", "--page-number", ...searchPageArgs(options)];
	if (options.ignoreCase) args.push("--ignore-case");
	if (options.literal) args.push("--fixed-strings");
	args.push("--", query, source);
	const result = await runCommand("pdfgrep", args, commandOptions(options, dirname(source)));
	if (result.code !== 0 && result.code !== 1) throw commandError(result, "pdfgrep");
	const matches = parseSearchOutput(result.stdout.toString("utf8"));
	return {
		matches,
		text: matches.map((match) => `page ${match.page}: ${match.text}`).join("\n"),
	};
}

async function outputPrefix(output: string, sourcePath: string): Promise<string> {
	const absolute = resolve(output);
	let isDirectory = output.endsWith("/") || output.endsWith("\\");
	try {
		isDirectory = (await stat(absolute)).isDirectory();
	} catch {
		// A not-yet-existing path is treated as a prefix unless it ends in a separator.
	}
	if (isDirectory) {
		await mkdir(absolute, { recursive: true });
		return join(absolute, basename(sourcePath, extname(sourcePath)));
	}
	await mkdir(dirname(absolute), { recursive: true });
	return absolute.replace(/\.png$/i, "");
}

function generatedPage(name: string, prefixName: string): number | undefined {
	if (name === `${prefixName}.png`) return 1;
	const match = name.match(new RegExp(`^${escapeRegex(prefixName)}-(\\d+)\\.png$`));
	return match ? Number(match[1]) : undefined;
}

async function generatedPages(directory: string, prefixName: string): Promise<string[]> {
	return (await readdir(directory))
		.map((name) => ({ name, page: generatedPage(name, prefixName) }))
		.filter((entry): entry is { name: string; page: number } => entry.page !== undefined)
		.sort((a, b) => a.page - b.page)
		.map(({ name }) => name);
}

async function clearGenerated(directory: string, prefixName: string): Promise<void> {
	for (const name of await generatedPages(directory, prefixName)) {
		await rm(join(directory, name), { force: true });
	}
}

export async function renderPdfPages(
	path: string,
	output: string,
	range: PageRange = {},
	options?: RunCommandOptions,
): Promise<string[]> {
	const source = resolve(path);
	const prefix = await outputPrefix(output, source);
	const directory = dirname(prefix);
	const prefixName = basename(prefix);
	const temporaryDirectory = await mkdtemp(join(directory, `.${prefixName}.`));
	const temporaryPrefix = join(temporaryDirectory, prefixName);
	try {
		const result = await runCommand(
			"pdftocairo",
			["-png", ...pageArgs(range), source, temporaryPrefix],
			commandOptions(options, dirname(source)),
		);
		if (result.code !== 0) throw commandError(result, "pdftocairo");

		const names = await generatedPages(temporaryDirectory, prefixName);
		if (names.length === 0) throw new Error("pdftocairo produced no PNG files");
		await clearGenerated(directory, prefixName);
		for (const name of names) {
			await rename(join(temporaryDirectory, name), join(directory, name));
		}
		return names.map((name) => join(directory, name));
	} finally {
		await rm(temporaryDirectory, { recursive: true, force: true });
	}
}
