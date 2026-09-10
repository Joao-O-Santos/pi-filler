import { mkdir, readdir, stat } from "node:fs/promises";
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

function pageArgs(range: PageRange): string[] {
	const args: string[] = [];
	if (range.firstPage !== undefined) args.push("-f", String(range.firstPage));
	if (range.lastPage !== undefined) args.push("-l", String(range.lastPage));
	return args;
}

function commandOptions(options: RunCommandOptions | undefined, cwd: string): RunCommandOptions {
	return { ...options, cwd: options?.cwd ?? cwd };
}

export async function readPdfText(
	path: string,
	range: PageRange = {},
	options?: RunCommandOptions,
): Promise<PdfTextResult> {
	const result = await runCommand(
		"pdftotext",
		["-layout", ...pageArgs(range), path, "-"],
		commandOptions(options, dirname(path)),
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
	options: RunCommandOptions & { ignoreCase?: boolean; literal?: boolean } = {},
): Promise<PdfSearchResult> {
	const pattern = options.literal ? escapeRegex(query) : query;
	const args = ["--color", "never", "--with-filename", "--page-number"];
	if (options.ignoreCase) args.push("--ignore-case");
	args.push("--", pattern, path);
	const result = await runCommand("pdfgrep", args, commandOptions(options, dirname(path)));
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

export async function renderPdfPages(
	path: string,
	output: string,
	range: PageRange = {},
	options?: RunCommandOptions,
): Promise<string[]> {
	const prefix = await outputPrefix(output, path);
	const args = ["-png", ...pageArgs(range), path, prefix];
	const result = await runCommand("pdftocairo", args, commandOptions(options, dirname(path)));
	if (result.code !== 0) throw commandError(result, "pdftocairo");

	const directory = dirname(prefix);
	const prefixName = basename(prefix);
	const generated = (await readdir(directory))
		.filter((name) => name.startsWith(`${prefixName}-`) && name.endsWith(".png"))
		.map((name) => join(directory, name));
	if (generated.length > 0) return generated.sort();

	const single = `${prefix}.png`;
	return [single];
}
