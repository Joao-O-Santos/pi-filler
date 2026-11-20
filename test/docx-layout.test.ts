import assert from "node:assert/strict";
import { link, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { zipSync } from "fflate";
import { type DocxLayout, inspectDocxLayout, patchDocxLayout } from "../src/docx-layout.js";
import { R, W } from "../src/docx-package.js";
import {
	directory,
	documentAt,
	drawing,
	enc,
	fixture,
	p,
	packageAt,
	row,
} from "./docx-edit-fixture.js";

async function setup(
	t: test.TestContext,
	body = p("Title") +
		p("Block lead") +
		drawing +
		p("First detail") +
		p("Final detail") +
		`<w:tbl>${row("Base")}</w:tbl>`,
) {
	const dir = await directory();
	t.after(() => rm(dir, { recursive: true, force: true }));
	const input = join(dir, "input.docx"),
		output = join(dir, "output.docx");
	const pkg = await fixture(
		input,
		body,
		{
			"word/media/plot.png": new Uint8Array([1, 2, 3]),
			"word/header1.xml": `<w:hdr xmlns:w="${W}">${p("Header")}</w:hdr>`,
		},
		`<Relationship Id="rIdImage" Type="${R}/image" Target="media/plot.png"/>`,
	);
	return { dir, input, output, pkg };
}

test("layout audit exposes stable indexes including image-only paragraphs and explicit/default distinction", async (t) => {
	const { input } = await setup(t);
	const audit = await inspectDocxLayout(input, 1, 2);
	assert.equal(audit.parts[0].paragraphCount, 7);
	assert.deepEqual(
		audit.parts[0].paragraphs.map((p) => p.index),
		[2, 3],
	);
	assert.equal(audit.parts[0].paragraphs[1].text, "");
	assert.equal(audit.parts[0].nextOffset, 3);
	assert.equal(audit.parts[0].tables[0].rows, 1);
	assert.match(audit.inheritance, /effective appearance is not resolved/);
	assert.equal(audit.sha256.length, 64);
	for (const [offset, limit] of [
		[-1, 1],
		[0, 0],
		[0, 101],
		[0.5, 1],
	])
		await assert.rejects(inspectDocxLayout(input, offset, limit), /offset\/limit/);
});

test("layout applies typography, terminates keep-chains, bounds images and clones formatted rows", async (t) => {
	const { input, output, pkg } = await setup(t);
	const result = await patchDocxLayout(input, output, {
		paragraphs: [
			{
				from: 1,
				to: 5,
				clearStyle: true,
				font: "Times New Roman",
				size: 11,
				line: 360,
				before: 160,
				after: 40,
				keepLines: true,
			},
			{ from: 1, alignment: "center", underline: true },
			{ from: 2, to: 4, keepNext: true, left: 340, hanging: 220 },
			{ from: 5, keepNext: false, pageBreakBefore: false },
		],
		cloneRows: [
			{
				table: 1,
				row: 1,
				count: 4,
				content: "keep",
				labels: ["Copy 1", "Copy 2", "Copy 3", "Copy 4"],
			},
		],
		imageBounds: { width: 15, height: 9 },
	});
	assert.equal(result.clonedRows, 4);
	assert.equal(result.resizedImages, 1);
	const doc = await documentAt(output);
	assert.match(doc, /Times New Roman/);
	assert.match(doc, /line="360"/);
	assert.match(doc, /hanging="220"/);
	assert.match(doc, /keepNext[^>]*val="0"/);
	assert.match(doc, /cx="2430000" cy="3240000"/);
	assert.equal((doc.match(/fill="000000"/g) ?? []).length, 5);
	assert.match(doc, />Copy 4<\/w:t>/);
	assert.equal((doc.match(/<w:gridSpan/g) ?? []).length, 5);
	assert.match(doc, /w:top="900"/);
	const out = await packageAt(output);
	for (const part of Object.keys(pkg).filter((name) => name !== "word/document.xml"))
		assert.deepEqual(out[part], pkg[part]);
});

test("math and textbox font runs stay untouched, explicit false overrides inherited formatting", async (t) => {
	const { input, output } = await setup(
		t,
		`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:beforeAutospacing="1"/><w:ind w:firstLine="100" w:leftChars="200"/></w:pPr><w:r><w:rPr><w:rStyle w:val="MockTextStyle"/></w:rPr><w:t>Text</w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><w:rPr><w:rStyle w:val="MockMathStyle"/><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>x</m:t></m:r></m:oMath><w:r><w:pict><w:txbxContent>${p("Textbox")}</w:txbxContent></w:pict></w:r></w:p>`,
	);
	await patchDocxLayout(input, output, {
		paragraphs: [
			{
				from: 1,
				font: "Times New Roman",
				size: 11,
				underline: false,
				keepNext: false,
				clearStyle: true,
				clearRunStyle: true,
				before: 10,
				left: 20,
				hanging: 30,
			},
		],
	});
	const doc = await documentAt(output);
	assert.match(doc, /Cambria Math/);
	assert.match(doc, /MockMathStyle/);
	assert.doesNotMatch(doc, /MockTextStyle/);
	assert.match(doc, /<w:txbxContent><w:p><w:r><w:t/);
	assert.doesNotMatch(doc, /pStyle|beforeAutospacing|firstLine|leftChars/);
	assert.match(doc, /val="none"/);
});

test("layout edits a named header without touching the document", async (t) => {
	const { input, output, pkg } = await setup(t);
	await patchDocxLayout(input, output, {
		part: "word/header1.xml",
		paragraphs: [{ from: 1, font: "Times New Roman", size: 11, alignment: "center" }],
	});
	const out = await packageAt(output);
	assert.deepEqual(out["word/document.xml"], pkg["word/document.xml"]);
	assert.match(new TextDecoder().decode(out["word/header1.xml"]), /Times New Roman/);
});

test("layout dry run and failed operations preserve both existing output and input", async (t) => {
	const { input, output } = await setup(t);
	const before = await readFile(input);
	await writeFile(output, "sentinel");
	const valid: DocxLayout = {
		paragraphs: [{ from: 1, to: 2, size: 11 }],
		cloneRows: [{ table: 1, row: 1, count: 1, content: "clear_text", labels: ["Copy 1"] }],
	};
	assert.equal((await patchDocxLayout(input, output, valid, true)).written, false);
	assert.equal(await readFile(output, "utf8"), "sentinel");
	const invalid = [
		{},
		{ part: "word/styles.xml", paragraphs: [{ from: 1, size: 11 }] },
		{ paragraphs: [{ from: 9, font: "TNR" }] },
		{ paragraphs: [{ from: 2, to: 1, size: 11 }] },
		{ paragraphs: [{ from: 1, size: 11.1 }] },
		{ paragraphs: [{ from: 1 }] },
		{ cloneRows: [{ table: 1, row: 1, count: 1, content: "keep", labels: ["x", "y"] }] },
		{ cloneRows: [{ table: 2, row: 1, count: 1, content: "keep" }] },
		{ cloneRows: [{ table: 1, row: 9, count: 1, content: "keep" }] },
		{
			cloneRows: [
				{ table: 1, row: 1, count: 1, content: "keep" },
				{ table: 1, row: 1, count: 1, content: "keep" },
			],
		},
	];
	for (const spec of invalid) {
		await assert.rejects(patchDocxLayout(input, output, spec as DocxLayout));
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
	await patchDocxLayout(input, output, valid);
	assert.match(await documentAt(output), />Copy 1<\/w:t>/);
	assert.deepEqual(await readFile(input), before);
});

test("cloning rejects vertical merges, fields, references, nested tables and duplicate IDs", async (t) => {
	const { input, output } = await setup(t);
	for (const unsafe of [
		'<w:vMerge w:val="restart"/>',
		"<w:hMerge/>",
		'<w:bookmarkStart w:id="1" w:name="b"/>',
		'<w:fldChar w:fldCharType="begin"/>',
		'<w:hyperlink r:id="rIdImage"/>',
		"<w:tbl/>",
	]) {
		await fixture(input, `<w:tbl>${row("Base").replace("</w:tr>", `${unsafe}</w:tr>`)}</w:tbl>`);
		await assert.rejects(
			patchDocxLayout(input, output, {
				cloneRows: [{ table: 1, row: 1, count: 1, content: "keep" }],
			}),
		);
	}
	for (const unsafe of ["<w:ins/>", "<w:sdt/>"]) {
		await fixture(input, p("X") + unsafe);
		await assert.rejects(
			patchDocxLayout(input, output, { paragraphs: [{ from: 1, size: 11 }] }),
			/unsupported/,
		);
	}
});

test("clones clear only requested text and never enlarge small images", async (t) => {
	const { input, output } = await setup(t);
	await patchDocxLayout(input, output, {
		cloneRows: [{ table: 1, row: 1, count: 1, content: "clear_text" }],
		imageBounds: { width: 100, height: 100 },
	});
	const doc = await documentAt(output);
	assert.equal((doc.match(/Cell content/g) ?? []).length, 1);
	assert.match(doc, /cx="5400000" cy="7200000"/);
});

test("layout rejects signed packages, malformed XML, aliases and supports intentional source replacement", async (t) => {
	const { input, output, pkg } = await setup(t);
	const spec = { paragraphs: [{ from: 1, size: 11 }] };
	await assert.rejects(patchDocxLayout(input, input, spec), /differ/);
	await link(input, output);
	await assert.rejects(patchDocxLayout(input, output, spec), /aliases/);
	await rm(output);
	await symlink(input, output);
	await assert.rejects(patchDocxLayout(input, output, spec), /aliases/);
	await rm(output);
	await patchDocxLayout(input, input, spec, false, true);
	assert.match(await documentAt(input), /val="22"/);
	await writeFile(input, zipSync({ ...pkg, "_xmlsignatures/sig1.xml": enc("<sig/>") }));
	await assert.rejects(patchDocxLayout(input, output, spec), /Signed/);
	await writeFile(input, zipSync({ ...pkg, "word/document.xml": enc("<bad") }));
	await assert.rejects(patchDocxLayout(input, output, spec));
});
