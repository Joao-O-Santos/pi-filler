// biome-ignore-all format: fixture-heavy regression tests preserve readable XML literals

import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";
import { readPptxText, searchPptx, writePptx } from "../src/pandoc.js";
import { inspectPptx, patchPptx } from "../src/pptx.js";

function xml(value: string): Uint8Array<ArrayBuffer> {
	return Uint8Array.from(new TextEncoder().encode(`<?xml version="1.0"?>${value}`));
}

async function pptxFixture(path: string): Promise<void> {
	const files = {
		"[Content_Types].xml": xml(
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
		),
		"ppt/presentation.xml": xml(
			'<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst/><p:sldIdLst><p:sldId id="1" r:id="slideOne"/><p:sldId id="2" r:id="slideTwo"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
		),
		"ppt/_rels/presentation.xml.rels": xml(
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slideOne" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="slideTwo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
		),
		"ppt/slides/slide1.xml": xml(
			'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
		),
		"ppt/slides/slide2.xml": xml(
			'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Other slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
		),
		"docProps/core.xml": xml(
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Original</dc:title><dc:creator>Alice</dc:creator></cp:coreProperties>',
		),
		"custom/preserved.bin": Uint8Array.from([8, 6, 7]),
	};
	await writeFile(path, zipSync(files));
}

function packagePart(path: string, part: string): Promise<Uint8Array> {
	return readFile(path).then((bytes) => unzipSync(new Uint8Array(bytes))[part]);
}

async function script(directory: string, name: string, body: string): Promise<void> {
	const path = join(directory, name);
	await writeFile(path, `#!/bin/sh\n${body}\n`);
	await chmod(path, 0o755);
}

test("inspects and patches PPTX slide XML transactionally", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pptx-patch-"));
	const input = join(directory, "input.pptx");
	const output = join(directory, "output.pptx");
	await pptxFixture(input);

	assert.deepEqual(await inspectPptx(input), {
		slideSize: { width: 12192000, height: 6858000 },
		slideCount: 2,
		metadata: { title: "Original", creator: "Alice" },
	});
	const result = await patchPptx(input, output, {
		slideSize: { width: 6858000, height: 12192000 },
		metadata: { title: "Updated" },
		replaceText: [{ find: "Hello world", replace: "Welcome", slide: 1 }],
	});
	assert.equal(result.written, true);
	assert.equal(result.replacementCount, 1);
	assert.deepEqual(result.changedParts, [
		"docProps/core.xml",
		"ppt/presentation.xml",
		"ppt/slides/slide1.xml",
	]);
	assert.equal((await inspectPptx(output)).metadata.title, "Updated");
	assert.match(new TextDecoder().decode(await packagePart(output, "ppt/slides/slide1.xml")), /Welcome/);
	assert.doesNotMatch(new TextDecoder().decode(await packagePart(output, "ppt/slides/slide1.xml")), /Hello|world/);
	assert.deepEqual([...await packagePart(output, "custom/preserved.bin")], [8, 6, 7]);
});

test("supports all-slide replacements, metadata clearing, and dry runs", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pptx-dry-run-"));
	const input = join(directory, "input.pptx");
	const output = join(directory, "output.pptx");
	await pptxFixture(input);

	const result = await patchPptx(input, output, {
		replaceText: [{ find: "slide", replace: "deck" }],
		clearCoreMetadata: true,
	}, true);
	assert.equal(result.written, false);
	assert.equal(result.replacementCount, 1);
	assert.equal(await exists(output), false);
	assert.deepEqual((await inspectPptx(input)).metadata, { title: "Original", creator: "Alice" });
});

test("rejects unsafe or ineffective PPTX patches", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pptx-errors-"));
	const input = join(directory, "input.pptx");
	const output = join(directory, "output.pptx");
	await pptxFixture(input);

	await assert.rejects(() => patchPptx(input, output, {}), /at least one change/);
	await assert.rejects(
		() => patchPptx(input, output, { replaceText: [{ find: "missing", replace: "x" }] }),
		/no matches/,
	);
	await assert.rejects(
		() => patchPptx(input, output, { replaceText: [{ find: "Hello", replace: "x", slide: 3 }] }),
		/out of range/,
	);
	await assert.rejects(
		() => patchPptx(input, output, { replaceText: [{ find: "", replace: "x" }] }),
		/non-empty find/,
	);
	await assert.rejects(() => patchPptx(input, input, { slideSize: { width: 1 } }), /differ from input/);
});

test("routes PPTX formatting and patch operations through filler", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pptx-route-"));
	const input = join(directory, "input.pptx");
	const output = join(directory, "output.pptx");
	await pptxFixture(input);

	const formatting = await executeFiller(
		{ format: "pptx", action: "read", view: "formatting", path: input },
		directory,
	);
	assert.match(formatting.text, /slideCount/);
	const patch = await executeFiller(
		{
			format: "pptx",
			action: "patch",
			path: input,
			output,
			pptx_patches: { slideSize: { width: 1 } },
			dry_run: true,
		},
		directory,
	);
	assert.equal(patch.details.written, false);
});

test("converts PPTX text through Pandoc and writes validated output", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pptx-pandoc-"));
	const input = join(directory, "input.pptx");
	const output = join(directory, "output.pptx");
	const source = join(directory, "slides.md");
	await pptxFixture(input);
	await writeFile(source, "# Slides\n\nContent\n");
	await script(
		directory,
		"pandoc",
		'if [ "$1" = "--from=pptx" ]; then printf "# Extracted\\n\\nNeedle\\n"; else cp "$PPTX_FIXTURE" "$4"; fi',
	);
	const env = { ...process.env, PATH: `${directory}:${process.env.PATH ?? ""}`, PPTX_FIXTURE: input };

	const read = await readPptxText(input, { env });
	assert.equal(read.text, "# Extracted\n\nNeedle\n");
	const search = await searchPptx(input, "needle", { env, ignoreCase: true });
	assert.equal(search.matchCount, 1);
	assert.equal(search.text, "line 3: Needle");
	await writePptx(source, output, { env, referencePptx: input });
	assert.equal((await inspectPptx(output)).slideCount, 2);
});

async function exists(path: string): Promise<boolean> {
	try {
		await stat(path);
		return true;
	} catch {
		return false;
	}
}
