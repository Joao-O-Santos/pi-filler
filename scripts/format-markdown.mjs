#!/usr/bin/env node

import { execFileSync, spawnSync } from "node:child_process";
import { readFile, rename, writeFile } from "node:fs/promises";
import { basename, dirname, join } from "node:path";

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
			"--columns=72",
			path,
		],
		{ encoding: "utf8", maxBuffer: 10 * 1024 * 1024 },
	);
	if (result.error) throw result.error;
	if (result.status !== 0) throw new Error(result.stderr.trim() || `Pandoc failed for ${path}`);
	return result.stdout;
}

const write = process.argv.includes("--write");
const changed = [];
for (const path of trackedMarkdown()) {
	const before = await readFile(path, "utf8");
	const after = formatMarkdown(path);
	if (before === after) continue;
	changed.push(path);
	if (!write) continue;
	const temporary = join(dirname(path), `.${basename(path)}.tmp`);
	await writeFile(temporary, after);
	await rename(temporary, path);
}

if (changed.length && !write) {
	console.error(`Markdown needs 72-column formatting:\n${changed.join("\n")}`);
	process.exitCode = 1;
} else if (changed.length) {
	console.log(`Formatted Markdown:\n${changed.join("\n")}`);
} else {
	console.log("Markdown matches 72-column formatting.");
}
