import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { zipSync } from "fflate";
import { inspectDocx, patchDocx } from "../src/docx.js";
import { executeFiller, normalizePath } from "../src/index.js";
import { renderPdfPages, searchPdf } from "../src/pdf.js";
import { runCommand } from "../src/process.js";

function xml(value: string): Uint8Array {
	return new TextEncoder().encode(`<?xml version="1.0"?>${value}`);
}

async function executable(path: string, content: string): Promise<void> {
	await writeFile(path, content);
	await chmod(path, 0o755);
}

async function docxFixture(path: string): Promise<void> {
	await writeFile(
		path,
		zipSync({
			"[Content_Types].xml": xml(
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
			),
			"word/document.xml": xml(
				'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
			),
			"docProps/core.xml": xml(
				'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title xml:lang="en">Original</dc:title></cp:coreProperties>',
			),
		}),
	);
}

async function exists(path: string): Promise<boolean> {
	try {
		await stat(path);
		return true;
	} catch {
		return false;
	}
}

test("PDF search honors page ranges and fixed-string mode", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-pdf-"));
	const log = join(directory, "pdfgrep-args");
	await executable(
		join(directory, "pdfgrep"),
		`#!/bin/sh\nprintf '%s\\n' "$@" > "$PDFGREP_LOG"\nprintf 'fixture.pdf:2:Needle found\\n'\n`,
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
	assert.match(args, /--with-filename/);
	assert.match(args, /--fixed-strings/);
	assert.match(args, /--page-range\n2-4/);
});

test("PDF search honors one-sided page ranges", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-range-"));
	const log = join(directory, "pdfgrep-args");
	await executable(
		join(directory, "pdfgrep"),
		`#!/bin/sh\nprintf '%s\\n' "$@" > "$PDFGREP_LOG"\nprintf 'fixture.pdf:3:Needle found\\n'\n`,
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

test("failed PDF rendering preserves previous output", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-render-"));
	await executable(join(directory, "pdftocairo"), "#!/bin/sh\nexit 2\n");
	const input = join(directory, "fixture.pdf");
	const stale = join(directory, "render-1.png");
	await writeFile(input, "fixture");
	await writeFile(stale, "previous");
	const env = { ...process.env, PATH: `${directory}:${process.env.PATH ?? ""}` };

	await assert.rejects(() => renderPdfPages(input, join(directory, "render"), {}, { env }));
	assert.equal(await exists(stale), true);
	assert.equal(await readFile(stale, "utf8"), "previous");
});

test("rejects ignored parameters and empty paths", async () => {
	assert.throws(() => normalizePath("@", process.cwd()), /Path must not be empty/);
	await assert.rejects(
		() =>
			executeFiller(
				{
					format: "pdf",
					action: "search",
					path: "fixture.pdf",
					query: "Needle",
					output: "ignored",
				},
				process.cwd(),
			),
		/PDF search does not accept output/,
	);
	await assert.rejects(
		() =>
			executeFiller(
				{
					format: "docx",
					action: "read",
					view: "text",
					path: "fixture.docx",
					first_page: 1,
				},
				process.cwd(),
			),
		/DOCX read does not accept first_page/,
	);
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

test("DOCX inspection rejects malformed XML", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-xml-"));
	const input = join(directory, "malformed.docx");
	await writeFile(
		input,
		zipSync({
			"[Content_Types].xml": xml(
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
			),
			"word/document.xml": xml('<w:document xmlns:w="x"><w:body></w:document>'),
		}),
	);
	await assert.rejects(() => inspectDocx(input), /Invalid XML in word\/document\.xml/);
});

test("DOCX patch rejects no-op and incompatible requests", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-patch-"));
	const input = join(directory, "input.docx");
	const output = join(directory, "output.docx");
	await docxFixture(input);

	await assert.rejects(() => patchDocx(input, output, {}), /at least one change/);
	await assert.rejects(
		() => patchDocx(input, output, { lineNumbering: { mode: "off", start: 1 } }),
		/mode=off cannot include/,
	);
});

test("orientation-only patch keeps page dimensions consistent", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-audit-orient-"));
	const input = join(directory, "input.docx");
	const output = join(directory, "output.docx");
	await docxFixture(input);

	await patchDocx(input, output, { page: { orientation: "landscape" } });
	const formatting = await inspectDocx(output);
	assert.equal(formatting.page.orientation, "landscape");
	assert.equal(formatting.page.width, 15840);
	assert.equal(formatting.page.height, 12240);
});
