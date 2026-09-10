import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { executeFiller } from "../src/index.js";
import { searchPdf } from "../src/pdf.js";
import { runCommand } from "../src/process.js";

async function executable(path: string, content: string): Promise<void> {
	await writeFile(path, content);
	await chmod(path, 0o755);
}

test("PDF search honors page ranges and fixed-string mode", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-pdf-"));
	const log = join(directory, "pdfgrep-args");
	await executable(
		join(directory, "pdfgrep"),
		`#!/bin/sh\nprintf '%s\\n' "$@" > "$PDFGREP_LOG"\nprintf '2:Needle found\\n'\n`,
	);
	const input = join(directory, "fixture.pdf");
	await writeFile(input, "fixture");
	const env = {
		...process.env,
		PATH: `${directory}:${process.env.PATH ?? ""}`,
		PDFGREP_LOG: log,
	};

	const result = await searchPdf(input, "Needle.*", {
		env,
		literal: true,
		firstPage: 2,
		lastPage: 4,
	});
	assert.deepEqual(result.matches, [{ page: 2, text: "Needle found" }]);
	const args = await readFile(log, "utf8");
	assert.match(args, /--no-filename/);
	assert.match(args, /--fixed-strings/);
	assert.match(args, /--page-range\n2-4/);
});

test("PDF search honors one-sided page ranges", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-range-"));
	const log = join(directory, "pdfgrep-args");
	await executable(
		join(directory, "pdfgrep"),
		`#!/bin/sh\nprintf '%s\\n' "$@" > "$PDFGREP_LOG"\nprintf '3:Needle found\\n'\n`,
	);
	const input = join(directory, "fixture.pdf");
	await writeFile(input, "fixture");
	const env = {
		...process.env,
		PATH: `${directory}:${process.env.PATH ?? ""}`,
		PDFGREP_LOG: log,
	};

	await searchPdf(input, "Needle", { env, firstPage: 3 });
	assert.match(await readFile(log, "utf8"), /--page-range\n3-2147483647/);
});

test("PDF search rejects formatting view", async () => {
	await assert.rejects(
		() =>
			executeFiller(
				{
					format: "pdf",
					action: "search",
					view: "formatting",
					path: "fixture.pdf",
					query: "Needle",
				},
				process.cwd(),
			),
		/PDF search requires view=text or no view/,
	);
});

test("command timeout terminates a process that ignores SIGTERM", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-timeout-"));
	await executable(
		join(directory, "stubborn"),
		`#!/usr/bin/env node\nprocess.on("SIGTERM", () => {});\nsetTimeout(() => process.exit(0), 3000);\n`,
	);
	const env = { ...process.env, PATH: `${directory}:${process.env.PATH ?? ""}` };
	const started = Date.now();
	await assert.rejects(() => runCommand("stubborn", [], { env, timeoutMs: 20 }), /timed out/);
	assert.ok(Date.now() - started < 1500, "timeout should not wait for ignored SIGTERM");
});
