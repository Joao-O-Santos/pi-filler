#!/usr/bin/env node

import { execFileSync, spawnSync } from "node:child_process";
import { readFile, rename, writeFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";

const WIDTH = 72;

function trackedMarkdown() {
	return execFileSync("git", ["ls-files", "*.md"], { encoding: "utf8" })
		.split("\n")
		.filter(Boolean);
}

function formatMarkdown(path) {
	const result = spawnSync(
		"pandoc",
		[
			"-f",
			"markdown",
			"-t",
			"markdown+pipe_tables-simple_tables-multiline_tables-grid_tables",
			"--wrap=auto",
			`--columns=${WIDTH}`,
			path,
		],
		{ encoding: "utf8", maxBuffer: 10 * 1024 * 1024 },
	);
	if (result.error) throw result.error;
	if (result.status !== 0) throw new Error(result.stderr.trim() || `Pandoc failed for ${path}`);
	return result.stdout;
}

function longProseLines(text) {
	const failures = [];
	let fenced = false;

	for (const [index, line] of text.split("\n").entries()) {
		if (/^\s{0,3}(```|~~~)/.test(line)) {
			fenced = !fenced;
			continue;
		}
		if (fenced || line.length <= WIDTH) continue;
		if (/^\s{0,3}#{1,6}\s/.test(line)) continue;
		if (/^\s*\[[^\]]+\]:\s+\S+/.test(line)) continue;
		if (/^\s*\S+:\/\/\S+\s*$/.test(line)) continue;
		if (/^\s*\|.*\|\s*$/.test(line)) continue;
		if (/^\s*</.test(line)) continue;
		failures.push(`${index + 1}:${line.length}`);
	}

	return failures;
}

const write = process.argv.includes("--write");
const failures = [];

for (const path of trackedMarkdown()) {
	const before = await readFile(path, "utf8");

	if (write) {
		const after = formatMarkdown(path);
		if (before === after) continue;
		const temporary = join(dirname(path), `.${basename(path)}.tmp`);
		await writeFile(temporary, after);
		await rename(temporary, path);
		continue;
	}

	const lines = longProseLines(before);
	if (lines.length) failures.push(`${path}: ${lines.join(", ")}`);
}

if (failures.length) {
	console.error(
		`Markdown prose exceeds ${WIDTH} characters (line:length):\n${failures.join("\n")}`,
	);
	process.exitCode = 1;
} else if (write) {
	console.log(`Formatted Markdown to ${WIDTH} columns.`);
} else {
	console.log(`Markdown prose fits ${WIDTH} columns.`);
}
