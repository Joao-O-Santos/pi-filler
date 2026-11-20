import assert from "node:assert/strict";
import { link, readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { zipSync } from "fflate";
import { composeDocx } from "../src/docx-compose.js";
import { hash, R, W } from "../src/docx-package.js";
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

async function setup(t: test.TestContext) {
	const dir = await directory();
	t.after(() => rm(dir, { recursive: true, force: true }));
	const shell = join(dir, "shell.docx"),
		donor = join(dir, "body.docx"),
		output = join(dir, "output.docx");
	const shellPkg = await fixture(
		shell,
		`${p("Cover")}<w:tbl>${row("Base")}</w:tbl>${p("OLD START")}${p("SECRET OLD BODY")}`,
		{
			"word/header1.xml": `<w:hdr xmlns:w="${W}">${p("Logo and PAGE fields untouched")}</w:hdr>`,
			"word/media/logo.png": new Uint8Array([1]),
		},
	);
	await fixture(
		donor,
		p("Donor preface") + p("New section") + drawing + p("First detail") + p("Final detail"),
		{ "word/media/plot.png": new Uint8Array([2, 3]) },
		`<Relationship Id="rIdImage" Type="${R}/image" Target="media/plot.png"/>`,
	);
	return { dir, shell, donor, output, shellPkg };
}

test("composition preserves synthetic shell parts, removes old body, and remaps image/style/IDs", async (t) => {
	const { shell, donor, output, shellPkg } = await setup(t);
	const result = await composeDocx(shell, output, {
		body_path: donor,
		replace_from: "OLD START",
		body_from: "New section",
		expected_sha256: hash(await readFile(shell)),
	});
	assert.equal(result.keptElements, 2);
	assert.equal(result.insertedElements, 4);
	assert.equal(result.importedRelationships, 1);
	assert.deepEqual(result.unresolvedDependencies, []);
	const doc = await documentAt(output);
	assert.match(doc, />Cover</);
	assert.match(doc, />New section</);
	assert.doesNotMatch(doc, /SECRET OLD BODY|OLD START|Donor preface/);
	assert.match(doc, /w:top="900"/);
	assert.match(doc, /r:embed="rIdFiller/);
	assert.match(doc, /FillerStyle/);
	const out = await packageAt(output);
	assert.deepEqual(out["word/header1.xml"], shellPkg["word/header1.xml"]);
	assert.deepEqual(out["word/media/logo.png"], shellPkg["word/media/logo.png"]);
	assert.deepEqual(out["word/media/filler1.png"], new Uint8Array([2, 3]));
	assert.match(new TextDecoder().decode(out["[Content_Types].xml"]), /\/word\/media\/filler1.png/);
});

test("composition dry-run, default body start, freshness, aliases and transactional refusals", async (t) => {
	const { shell, donor, output } = await setup(t);
	await writeFile(output, "sentinel");
	const spec = { body_path: donor, replace_from: "OLD START" };
	assert.equal((await composeDocx(shell, output, spec, true)).insertedElements, 5);
	assert.equal(await readFile(output, "utf8"), "sentinel");
	for (const change of [
		{ replace_from: "absent" },
		{ body_from: "absent" },
		{ expected_sha256: "0".repeat(64) },
		{ extra: true },
	]) {
		await assert.rejects(composeDocx(shell, output, { ...spec, ...change }));
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
	await assert.rejects(composeDocx(shell, shell, spec), /differ/);
	await assert.rejects(composeDocx(shell, donor, spec), /differ/);
	await rm(output);
	await link(shell, output);
	await assert.rejects(composeDocx(shell, output, spec), /aliases/);
	await rm(output);
	await fixture(shell, p("OLD START") + p("OLD START"));
	await assert.rejects(composeDocx(shell, output, spec), /exactly one/);
});

test("composition refuses unsupported dependencies rather than silently dropping them", async (t) => {
	const { shell, donor, output } = await setup(t);
	const spec = { body_path: donor, replace_from: "OLD START" };
	const bad = [
		'<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>',
		"<w:p><w:ins/></w:p>",
		"<w:p><w:r><w:fldChar/></w:r></w:p>",
		"<w:p><w:pPr><w:sectPr/></w:pPr></w:p>",
		'<w:p><w:hyperlink w:anchor="b"/></w:p>',
		"<w:p><w:r><w:pict/></w:r></w:p>",
		'<w:altChunk r:id="chunk"/>',
	];
	for (const body of bad) {
		await fixture(donor, p("New") + body);
		await writeFile(output, "sentinel");
		await assert.rejects(composeDocx(shell, output, spec));
		assert.equal(await readFile(output, "utf8"), "sentinel");
	}
	await fixture(donor, p("A") + drawing);
	await assert.rejects(composeDocx(shell, output, spec), /relationship ID/);
	await fixture(donor, p("A"), { "_xmlsignatures/sig.xml": "<sig/>" });
	await assert.rejects(composeDocx(shell, output, spec), /Signed/);
});

test("composition imports safe external hyperlinks and refuses linked images, charts and local links", async (t) => {
	const { shell, donor, output } = await setup(t);
	const spec = { body_path: donor, replace_from: "OLD START" };
	const body = `<w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>`;
	await fixture(
		donor,
		body,
		{},
		`<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.org" TargetMode="External"/>`,
	);
	await composeDocx(shell, output, spec);
	assert.match(
		new TextDecoder().decode((await packageAt(output))["word/_rels/document.xml.rels"]),
		/https:\/\/example.org/,
	);
	for (const rel of [
		`<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="file:///etc/passwd" TargetMode="External"/>`,
		`<Relationship Id="rIdLink" Type="${R}/image" Target="https://example.org/x.png" TargetMode="External"/>`,
		`<Relationship Id="rIdLink" Type="${R}/chart" Target="chart.xml"/>`,
	]) {
		await fixture(donor, body, { "word/chart.xml": "<chart/>" }, rel);
		await assert.rejects(composeDocx(shell, output, spec));
	}
});

test("composition remaps numbered paragraphs and style links without changing shell styles", async (t) => {
	const { shell, donor, output } = await setup(t);
	const numbering = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Normal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
	await fixture(
		donor,
		`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p>`,
		{ "word/numbering.xml": numbering },
		`<Relationship Id="rIdNum" Type="${R}/numbering" Target="numbering.xml"/>`,
	);
	const result = await composeDocx(shell, output, { body_path: donor, replace_from: "OLD START" });
	assert.equal(result.importedNumbering, 1);
	const pkg = await packageAt(output);
	const merged = new TextDecoder().decode(pkg["word/styles.xml"]);
	assert.match(merged, /w:styleId="Normal"/);
	assert.match(merged, /w:styleId="FillerStyle/);
	assert.ok(Object.keys(pkg).some((name) => name.includes("fillerNumbering")));
});

test("composition handles namespace collisions and bookmark/drawing IDs independently", async (t) => {
	const { shell, donor, output } = await setup(t);
	await fixture(
		donor,
		`<w:p><w:bookmarkStart w:id="0" w:name="old"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`,
	);
	const pkg = await packageAt(donor);
	pkg["word/document.xml"] = enc(
		new TextDecoder()
			.decode(pkg["word/document.xml"])
			.replaceAll("w:", "x:")
			.replace("xmlns:w=", "xmlns:x="),
	);
	await writeFile(donor, zipSync(pkg));
	await composeDocx(shell, output, { body_path: donor, replace_from: "OLD START" });
	const doc = await documentAt(output);
	assert.match(doc, /xmlns:x=/);
	assert.match(doc, /FillerBookmark1/);
	assert.match(doc, /<x:bookmarkEnd x:id="1"/);
});
