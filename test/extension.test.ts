import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { unzipSync, zipSync } from "fflate";
import { inspectDocx, patchDocx } from "../src/docx.js";
import extension, { executeFiller, normalizePath } from "../src/index.js";
import { readDocxText, searchDocx, writeDocx } from "../src/pandoc.js";
import { readPdfText, renderPdfPages, searchPdf } from "../src/pdf.js";
import { MissingExecutableError, runCommand } from "../src/process.js";

async function writeFixture(path: string, wordPrefix = "w"): Promise<void> {
	const xml = (value: string) => new TextEncoder().encode(`<?xml version="1.0"?>${value}`);
	const w = wordPrefix;
	const files = {
		"[Content_Types].xml": xml(
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
		),
		"word/document.xml": xml(
			`<${w}:document xmlns:${w}="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><${w}:body><${w}:p/><${w}:sectPr><${w}:pgSz ${w}:w="12240" ${w}:h="15840"/><${w}:pgMar ${w}:top="1440" ${w}:bottom="1440" ${w}:left="1440" ${w}:right="1440"/><${w}:lnNumType ${w}:restart="continuous"/><${w}:pgNumType ${w}:start="1" ${w}:fmt="decimal"/></${w}:sectPr></${w}:body></${w}:document>`,
		),
		"word/_rels/document.xml.rels": xml(
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/image.png" Type="image"/></Relationships>',
		),
		"word/media/image.png": new Uint8Array([1, 2, 3]),
		"word/comments.xml": xml(
			`<${w}:comments xmlns:${w}="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><${w}:comment ${w}:id="0" ${w}:author="Alice"/></${w}:comments>`,
		),
		"docProps/core.xml": xml(
			'<props:coreProperties xmlns:props="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:d="http://purl.org/dc/elements/1.1/"><d:title>Original</d:title><d:creator>Alice</d:creator><props:lastModifiedBy>Bob</props:lastModifiedBy></props:coreProperties>',
		),
		"custom/unknown.bin": new Uint8Array([9, 8, 7]),
	};
	await writeFile(path, zipSync(files));
}

async function fakeCommands(): Promise<{ directory: string; env: NodeJS.ProcessEnv }> {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-test-"));
	const script = async (name: string, body: string) => {
		const path = join(directory, name);
		await writeFile(path, `#!/bin/sh\n${body}\n`);
		await chmod(path, 0o755);
	};
	await script("pdftotext", 'printf "%s\\n" "$@" > "$PDFTOTEXT_LOG"\nprintf "first page\\nsecond page\\n"');
	await script("pdfgrep", 'printf "%s\\n" "$@" > "$PDFGREP_LOG"\nprintf "fixture.pdf:2:Needle found\\n"');
	await script(
		"pdftocairo",
		'printf "%s\\n" "$@" > "$PDFTOCAIRO_LOG"\nfor arg do prefix="$arg"; done\ntouch "$prefix-10.png" "$prefix-2.png"',
	);
	await script(
		"pandoc",
		'printf "%s\\n" "$@" > "$PANDOC_LOG"\n' +
			'if [ "$PANDOC_FAIL" = "1" ]; then echo "pandoc failed" >&2; exit 2; fi\n' +
			'if [ "$1" = "--from=gfm" ]; then\n' +
			'  previous=""\n' +
			'  for arg do if [ "$previous" = "--output" ]; then output="$arg"; fi; previous="$arg"; done\n' +
			'  if [ "$PANDOC_INVALID" = "1" ]; then printf "not docx" > "$output"; else cp "$PANDOC_DOCX_FIXTURE" "$output"; fi\n' +
			'elif [ "$PANDOC_LONG" = "1" ]; then\n' +
			'  head -c 60000 /dev/zero | tr "\\000" x\n' +
			'  printf "\\nLate needle\\n"\n' +
			"else\n" +
			'  printf "# Heading\\nBody text\\nBody match\\n"\n' +
			"fi",
	);
	await script("slow-command", "sleep 1");
	await script("large-command", "head -c 200000 /dev/zero");
	const fixture = join(directory, "pandoc-output.docx");
	await writeFixture(fixture);
	return {
		directory,
		env: {
			...process.env,
			PATH: `${directory}:${process.env.PATH ?? ""}`,
			PANDOC_LOG: join(directory, "pandoc-args"),
			PANDOC_DOCX_FIXTURE: fixture,
			PDFTOTEXT_LOG: join(directory, "pdftotext-args"),
			PDFGREP_LOG: join(directory, "pdfgrep-args"),
			PDFTOCAIRO_LOG: join(directory, "pdftocairo-args"),
		},
	};
}

test("extension registers one filler tool", () => {
	let registered: { name?: string } | undefined;
	extension({
		registerTool(definition: { name?: string }) {
			registered = definition;
		},
	} as never);
	assert.equal(registered?.name, "filler");
});

test("reports missing executables clearly", async () => {
	await assert.rejects(
		() => runCommand("pi-filler-command-that-does-not-exist", []),
		(error: unknown) =>
			error instanceof MissingExecutableError && /available on PATH/.test(error.message),
	);
});

test("bounds and times out external commands", async () => {
	const { env } = await fakeCommands();
	await assert.rejects(() => runCommand("large-command", [], { env, maxBufferBytes: 1024 }), /exceeded/);
	await assert.rejects(() => runCommand("slow-command", [], { env, timeoutMs: 10 }), /timed out/);
	const controller = new AbortController();
	controller.abort();
	await assert.rejects(() => runCommand("slow-command", [], { env, signal: controller.signal }), /aborted/);
});

test("normalizes relative and @ paths", () => {
	assert.equal(normalizePath("@docs/input.pdf", "/work/project"), "/work/project/docs/input.pdf");
	assert.equal(normalizePath("/tmp/input.pdf", "/work/project"), "/tmp/input.pdf");
});

test("runs deterministic PDF commands", async () => {
	const { directory, env } = await fakeCommands();
	const input = join(directory, "fixture.pdf");
	await writeFile(input, "not a real PDF");

	const text = await readPdfText(input, { firstPage: 2, lastPage: 2 }, { env });
	assert.equal(text.text, "first page\nsecond page\n");
	assert.match(await readFile(env.PDFTOTEXT_LOG ?? "", "utf8"), /-f\n2\n-l\n2/);

	const search = await searchPdf(input, "Needle", { env, literal: true, ignoreCase: true });
	assert.deepEqual(search.matches, [{ page: 2, text: "Needle found" }]);
	assert.equal(search.text, "page 2: Needle found");
	const searchArgs = await readFile(env.PDFGREP_LOG ?? "", "utf8");
	assert.match(searchArgs, /--page-number/);
	assert.match(searchArgs, /--ignore-case/);

	const stale = join(directory, "render-99.png");
	await writeFile(stale, "stale");
	const images = await renderPdfPages(input, join(directory, "render"), {}, { env });
	assert.deepEqual(images, [join(directory, "render-2.png"), join(directory, "render-10.png")]);
	assert.equal(await statSafe(stale), false);
	assert.match(await readFile(env.PDFTOCAIRO_LOG ?? "", "utf8"), /-png/);
});

test("reads and searches DOCX through full Pandoc output", async () => {
	const { directory, env } = await fakeCommands();
	const input = join(directory, "fixture.docx");
	await writeFile(input, "fake docx");

	const text = await readDocxText(input, { env });
	assert.equal(text.text, "# Heading\nBody text\nBody match\n");
	const search = await searchDocx(input, "match", { env });
	assert.deepEqual(search.matches, [{ line: 3, text: "Body match" }]);
	const late = await searchDocx(input, "needle", { env: { ...env, PANDOC_LONG: "1" }, ignoreCase: true });
	assert.equal(late.matches.length, 1);
	assert.equal(late.matches[0]?.text, "Late needle");
});

test("writes and validates transactional DOCX output", async () => {
	const { directory, env } = await fakeCommands();
	const source = join(directory, "source.md");
	const output = join(directory, "nested", "result.docx");
	const reference = join(directory, "reference.docx");
	const markdown = "# Source\n\nContent\n";
	await writeFile(source, markdown);
	await writeFile(reference, "reference");

	await writeDocx(source, output, { env, referenceDocx: reference });
	assert.equal(await readFile(source, "utf8"), markdown);
	assert.equal((await inspectDocx(output)).metadata.title, "Original");
	const args = await readFile(env.PANDOC_LOG ?? "", "utf8");
	assert.match(args, /--from=gfm/);
	assert.match(args, /--reference-doc/);
	assert.match(args, /source\.md/);

	const invalid = join(directory, "invalid.docx");
	await assert.rejects(
		() => writeDocx(source, invalid, { env: { ...env, PANDOC_INVALID: "1" } }),
		/Invalid DOCX ZIP/,
	);
	assert.equal(await statSafe(invalid), false);
});

test("reports Pandoc failures and rejects DOCX input/output identity", async () => {
	const { directory, env } = await fakeCommands();
	const source = join(directory, "source.md");
	await writeFile(source, "content");
	await assert.rejects(
		() => writeDocx(source, join(directory, "out.docx"), { env: { ...env, PANDOC_FAIL: "1" } }),
		/pandoc failed/,
	);
	await assert.rejects(
		() =>
			executeFiller({ format: "docx", action: "write", path: source, output: source }, directory),
		/DOCX write output must differ from input/,
	);
});

test("inspects and patches DOCX OOXML while preserving unknown parts", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-ooxml-"));
	const input = join(directory, "input.docx");
	const output = join(directory, "output.docx");
	await writeFixture(input, "x");
	const before = await inspectDocx(input);
	assert.equal(before.sectionCount, 1);
	assert.equal(before.page.width, 12240);
	assert.equal(before.lineNumbering.mode, "continuous");
	assert.equal(before.comments, 1);
	assert.equal(before.metadata.title, "Original");
	const dry = await patchDocx(
		input,
		output,
		{
			page: { orientation: "landscape", width: 15840, height: 12240 },
			margins: { top: 720, bottom: 720 },
			lineNumbering: { mode: "newSection", start: 1, count_by: 5 },
			pageNumbering: { start: 3, format: "upperRoman" },
			clearCoreMetadata: true,
			metadata: { title: "Changed" },
		},
		true,
	);
	assert.equal(dry.written, false);
	assert.ok(dry.changedParts.includes("word/document.xml"));
	assert.ok(dry.changedParts.includes("docProps/core.xml"));
	assert.equal(await statSafe(output), false);
	const result = await patchDocx(input, output, {
		page: { orientation: "landscape", width: 15840, height: 12240 },
		margins: { top: 720, bottom: 720 },
		lineNumbering: { mode: "newSection", start: 1, count_by: 5 },
		pageNumbering: { start: 3, format: "upperRoman" },
		clearCoreMetadata: true,
		metadata: { title: "Changed" },
	});
	assert.equal(result.written, true);
	assert.deepEqual(await unzipUnknown(output), [9, 8, 7]);
	const after = await inspectDocx(output);
	assert.equal(after.page.orientation, "landscape");
	assert.equal(after.margins.top, 720);
	assert.equal(after.lineNumbering.mode, "newSection");
	assert.equal(after.lineNumbering.count_by, 5);
	assert.equal(after.pageNumbering.format, "upperRoman");
	assert.equal(after.metadata.title, "Changed");
	assert.equal(after.metadata.creator, "");
	assert.equal(after.metadata.lastModifiedBy, "");
});

test("validates every internal relationship target", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-rels-"));
	const input = join(directory, "broken.docx");
	await writeFixture(input);
	const files = unzipSync(new Uint8Array(await readFile(input)));
	files["word/_rels/header1.xml.rels"] = new TextEncoder().encode(
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/missing.png" Type="image"/></Relationships>',
	);
	await writeFile(input, zipSync(files));
	await assert.rejects(() => inspectDocx(input), /relationship target is missing/);
});

test("rejects invalid packages and unsupported combinations", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-invalid-"));
	await writeFile(join(directory, "bad.docx"), "not zip");
	await assert.rejects(() => inspectDocx(join(directory, "bad.docx")), /Invalid DOCX ZIP/);

	const pdf = join(directory, "fixture.pdf");
	await writeFile(pdf, "not a real PDF");
	await assert.rejects(
		() => executeFiller({ format: "pdf", action: "write", path: pdf }, directory),
		/PDF write and patch operations are not supported/,
	);
	await assert.rejects(
		() => executeFiller({ format: "pdf", action: "read", path: pdf }, directory),
		/PDF read requires view=text or view=image/,
	);
});

async function statSafe(path: string): Promise<boolean> {
	try {
		await stat(path);
		return true;
	} catch {
		return false;
	}
}

async function unzipUnknown(path: string): Promise<number[]> {
	return Array.from(unzipSync(new Uint8Array(await readFile(path)))["custom/unknown.bin"]);
}
