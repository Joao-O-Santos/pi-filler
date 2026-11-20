import assert from "node:assert/strict";
import { readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { zipSync } from "fflate";
import { composeDocx } from "../src/docx-compose.js";
import { inspectDocxLayout, patchDocxLayout } from "../src/docx-layout.js";
import { all, CT, children, get, is, load, parse, R, REL, text, W } from "../src/docx-package.js";
import {
	directory,
	documentAt,
	enc,
	fixture,
	p,
	packageAt,
	row,
	styles,
} from "./docx-edit-fixture.js";

async function setup(t: test.TestContext) {
	const dir = await directory();
	t.after(() => rm(dir, { recursive: true, force: true }));
	const shell = join(dir, "shell.docx"),
		donor = join(dir, "body.docx"),
		output = join(dir, "output.docx");
	await fixture(shell, p("Cover") + p("REPLACE") + p("Old body"));
	await fixture(donor, p("New body"));
	await writeFile(output, "sentinel");
	return { shell, donor, output };
}

test("layout source indexes remain original when cloning more than one row", async (t) => {
	const { shell, output } = await setup(t);
	await fixture(shell, `<w:tbl>${row("Row A")}${row("Row B")}</w:tbl>`);
	await patchDocxLayout(shell, output, {
		cloneRows: [
			{ table: 1, row: 1, count: 1, content: "keep", labels: ["Copy A"] },
			{ table: 1, row: 2, count: 1, content: "keep", labels: ["Copy B"] },
		],
	});
	const { pkg } = await load(output);
	const root = parse(pkg["word/document.xml"], "document");
	const table = all(root).find((n) => is(n, W, "tbl"));
	assert.ok(table);
	assert.deepEqual(
		children(table, W, "tr").map((n) => text(children(n, W, "tc")[0])),
		["Row A", "Copy A", "Row B", "Copy B"],
	);
});

test("invalid XML strings, UTF-8 and references are refused before output publication, including dry runs", async (t) => {
	const { shell, output } = await setup(t);
	for (const invalid of ["\0", "\ud800", "\uffff"]) {
		await assert.rejects(
			patchDocxLayout(shell, output, { paragraphs: [{ from: 1, font: invalid }] }, true),
			/XML/,
		);
		await fixture(shell, `<w:tbl>${row("Row B")}</w:tbl>`);
		await assert.rejects(
			patchDocxLayout(
				shell,
				output,
				{ cloneRows: [{ table: 1, row: 1, count: 1, content: "keep", labels: [invalid] }] },
				true,
			),
			/XML/,
		);
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
	const base = await fixture(shell, p("X"));
	for (const part of [
		enc(`<w:document xmlns:w="${W}"><w:body>${p("&#0;")}</w:body></w:document>`),
		new Uint8Array([0xff, 0xfe]),
	]) {
		await writeFile(shell, zipSync({ ...base, "word/document.xml": part }));
		await assert.rejects(
			patchDocxLayout(shell, output, { paragraphs: [{ from: 1, size: 11 }] }, true),
		);
	}
});

test("XML-sensitive labels, character-reference CR, numeric-looking text, comments and spaces survive editing", async (t) => {
	const { shell, output } = await setup(t);
	await fixture(
		shell,
		`${p("0001.20") + p("a&#13;b")}<!--retained--><w:tbl>${row("Row B")}</w:tbl>`,
	);
	await patchDocxLayout(shell, output, {
		paragraphs: [{ from: 1, to: 2, size: 11 }],
		cloneRows: [{ table: 1, row: 1, count: 1, content: "keep", labels: ["  A & <B>\r C  "] }],
	});
	const doc = await documentAt(output);
	assert.match(doc, /0001.20/);
	assert.match(doc, /<!--retained-->/);
	assert.match(doc, /a&#13;b/);
	assert.match(doc, /xml:space="preserve"/);
	const root = parse((await packageAt(output))["word/document.xml"], "document");
	assert.ok(text(root).includes("  A & <B>\r C  "));
});

test("structure audits byte-bound large explicit records, paginate, select header stories and identify reference types", async (t) => {
	const { shell } = await setup(t);
	const attrs = Array.from({ length: 8 }, (_, i) => `w:custom${i}="${"x".repeat(200)}"`).join(" ");
	const hugeProps = Array.from({ length: 3 }, (_, i) => `<w:custom${i} ${attrs}/>`).join("");
	const body = Array.from(
		{ length: 100 },
		() =>
			`<w:p><w:pPr>${hugeProps}</w:pPr>${Array.from({ length: 9 }, () => `<w:r><w:rPr>${hugeProps}</w:rPr><w:t>x</w:t></w:r>`).join("")}</w:p>`,
	).join("");
	const pkg = await fixture(
		shell,
		body,
		{ "word/header1.xml": `<w:hdr xmlns:w="${W}">${p("First-page header")}</w:hdr>` },
		`<Relationship Id="header" Type="${R}/header" Target="header1.xml"/>`,
	);
	pkg["word/document.xml"] = enc(
		new TextDecoder()
			.decode(pkg["word/document.xml"])
			.replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="first" r:id="header"/>'),
	);
	await writeFile(shell, zipSync(pkg));
	const result = await inspectDocxLayout(shell, 0, 100);
	assert.ok(Buffer.byteLength(JSON.stringify(result)) < 50000);
	assert.ok(result.parts[0].nextOffset !== null);
	assert.deepEqual(result.sectionReferences, [
		{ section: 1, kind: "header", type: "first", part: "word/header1.xml" },
	]);
	const header = await inspectDocxLayout(shell, 0, 1, "word/header1.xml");
	assert.equal(header.parts[0].paragraphs[0].text, "First-page header");
	await assert.rejects(inspectDocxLayout(shell, 0, 1, "word/header9.xml"), /Unknown/);
});

test("composition refuses boundary-crossing bookmarks and removed-anchor links; PAGE fields in retained cover are preserved", async (t) => {
	const { shell, donor, output } = await setup(t);
	const spec = { body_path: donor, replace_from: "REPLACE" };
	for (const body of [
		`<w:p><w:bookmarkStart w:id="1" w:name="b"/></w:p>${p("REPLACE")}<w:p><w:bookmarkEnd w:id="1"/></w:p>`,
		`<w:p><w:hyperlink w:anchor="b"/></w:p>${p("REPLACE")}<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/></w:p>`,
		`<w:p><w:r><w:instrText> REF old </w:instrText></w:r></w:p>${p("REPLACE")}`,
	]) {
		await fixture(shell, body);
		await assert.rejects(composeDocx(shell, output, spec));
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
	await fixture(shell, `<w:p><w:r><w:instrText> PAGE </w:instrText></w:r></w:p>${p("REPLACE")}`);
	await composeDocx(shell, output, spec);
	assert.match(await documentAt(output), / PAGE /);
});

test("composition refuses missing/duplicate styles only if actually used, and differing themes explicitly", async (t) => {
	const { shell, donor, output } = await setup(t);
	const spec = { body_path: donor, replace_from: "REPLACE" };
	const body =
		'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>';
	for (const value of [
		styles.replace('w:styleId="Heading1"', 'w:styleId="Other"'),
		styles.replace("</w:styles>", '<w:style w:type="paragraph" w:styleId="Heading1"/></w:styles>'),
	]) {
		await fixture(donor, body, { "word/styles.xml": value });
		await assert.rejects(composeDocx(shell, output, spec), /style ID/);
	}
	await fixture(donor, p("New"), {
		"word/styles.xml": styles.replace('w:styleId="Heading1"', 'w:styleId="Normal"'),
	});
	await assert.rejects(composeDocx(shell, output, spec), /style/);
	await fixture(donor, body, {
		"word/styles.xml": styles.replace('w:val="48"', 'w:val="48" w:theme="1"'),
	});
	await assert.rejects(composeDocx(shell, output, spec), /themes/);
});

test("composition rejects malformed bookmark pairing, numbering and relationship escapes", async (t) => {
	const { shell, donor, output } = await setup(t);
	const spec = { body_path: donor, replace_from: "REPLACE" };
	for (const markers of [
		'<w:bookmarkStart w:id="1" w:name="b"/>',
		'<w:bookmarkEnd w:id="1"/>',
		'<w:bookmarkStart w:id="1"/><w:bookmarkStart w:id="1"/>',
		'<w:bookmarkStart w:id="1"/><w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="1"/>',
	]) {
		await fixture(donor, `<w:p>${markers}</w:p>`);
		await assert.rejects(composeDocx(shell, output, spec), /bookmark/i);
	}
	await fixture(donor, '<w:p><w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr></w:p>');
	await assert.rejects(composeDocx(shell, output, spec), /numbering/);
	await fixture(
		donor,
		p("X"),
		{},
		`<Relationship Id="bad" Type="${R}/image" Target="../../secret.png"/>`,
	);
	await assert.rejects(composeDocx(shell, output, spec), /escapes/);
	assert.equal(await readFile(output, "utf8"), "sentinel");
});

test("structure defaults resolve through the styles relationship, not a guessed filename", async (t) => {
	const { shell } = await setup(t);
	const pkg = await packageAt(shell);
	pkg["word/_rels/document.xml.rels"] = enc(
		new TextDecoder()
			.decode(pkg["word/_rels/document.xml.rels"])
			.replace('Target="styles.xml"', 'Target="custom/styles.xml"'),
	);
	pkg["[Content_Types].xml"] = enc(
		new TextDecoder()
			.decode(pkg["[Content_Types].xml"])
			.replace("/word/styles.xml", "/word/custom/styles.xml"),
	);
	pkg["word/custom/styles.xml"] = pkg["word/styles.xml"];
	delete pkg["word/styles.xml"];
	await writeFile(shell, zipSync(pkg));
	const result = await inspectDocxLayout(shell);
	assert.match(JSON.stringify(result.defaults), /Times New Roman/);
});

test("page-break removal keeps line/column breaks; twip overrides keep foreign attributes", async (t) => {
	const { shell, output } = await setup(t);
	await fixture(
		shell,
		`<w:p xmlns:c="urn:custom"><w:pPr><w:wordWrap/><w:spacing w:beforeLines="100" w:afterLines="100" c:beforeLines="unchanged"/><w:jc w:val="left"/></w:pPr><w:r><w:br w:type="page"/><w:br w:type="column"/><w:br/><w:t>X</w:t></w:r></w:p>`,
	);
	await patchDocxLayout(shell, output, {
		paragraphs: [{ from: 1, clearPageBreaks: true, before: 10, after: 20, keepNext: true }],
	});
	const doc = await documentAt(output);
	assert.doesNotMatch(doc, /w:type="page"|w:beforeLines|w:afterLines/);
	assert.match(doc, /w:type="column"/);
	assert.match(doc, /c:beforeLines="unchanged"/);
	assert.ok(doc.indexOf("keepNext") < doc.indexOf("wordWrap"));
});

test("composition avoids preexisting bookmark names and rejects reversed bookmark ranges", async (t) => {
	const { shell, donor, output } = await setup(t);
	await fixture(
		shell,
		`<w:p><w:bookmarkStart w:id="10" w:name="FillerBookmark1"/><w:r><w:t>Cover</w:t></w:r><w:bookmarkEnd w:id="10"/></w:p>${p("REPLACE")}`,
	);
	await fixture(
		donor,
		'<w:p><w:bookmarkStart w:id="0" w:name="B"/><w:r><w:t>New</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>',
	);
	await composeDocx(shell, output, { body_path: donor, replace_from: "REPLACE" });
	const root = parse((await packageAt(output))["word/document.xml"], "document");
	const names = all(root)
		.filter((n) => is(n, W, "bookmarkStart"))
		.map((n) => get(n, W, "name"));
	assert.deepEqual(names, ["FillerBookmark1", "FillerBookmark2"]);
	await fixture(
		donor,
		'<w:p><w:bookmarkEnd w:id="0"/><w:bookmarkStart w:id="0" w:name="B"/></w:p>',
	);
	await assert.rejects(
		composeDocx(shell, output, { body_path: donor, replace_from: "REPLACE" }),
		/bookmark/,
	);
});

test("signatures, wrong namespaces, ambiguous IDs and doctypes are refused transactionally", async (t) => {
	const { shell, output } = await setup(t);
	const base = await packageAt(shell);
	const changes: Record<string, Uint8Array>[] = [
		{
			"[Content_Types].xml": enc(
				`<Types xmlns="${CT}"><Override PartName="/sig.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/></Types>`,
			),
		},
		{
			"word/_rels/document.xml.rels": enc(
				`<Relationships xmlns="${REL}"><Relationship Id="sig" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="http://invalid" TargetMode="External"/></Relationships>`,
			),
		},
		{ "word/_rels/document.xml.rels": enc(`<Relationships xmlns="foreign"/>`) },
		{ "[Content_Types].xml": enc(`<Types xmlns="foreign"/>`) },
		{
			"word/_rels/document.xml.rels": enc(
				`<Relationships xmlns="${REL}"><Relationship Id="a" Target="styles.xml"/><Relationship Id="a" Target="styles.xml"/></Relationships>`,
			),
		},
		{
			"word/document.xml": enc(
				`<!DOCTYPE x [<!ENTITY data SYSTEM "file:///etc/passwd">]><w:document xmlns:w="${W}"><w:body/></w:document>`,
			),
		},
	];
	for (const change of changes) {
		await writeFile(shell, zipSync({ ...base, ...change }));
		await assert.rejects(
			patchDocxLayout(shell, output, { paragraphs: [{ from: 1, size: 11 }] }, true),
		);
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
});
