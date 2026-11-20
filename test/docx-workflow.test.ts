import assert from "node:assert/strict";
import { readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import { deflateSync } from "node:zlib";
import { zipSync } from "fflate";
import { composeDocx } from "../src/docx-compose.js";
import { inspectDocxLayout, patchDocxLayout } from "../src/docx-layout.js";
import {
	CT,
	children,
	elements,
	hash,
	is,
	parse,
	R,
	REL,
	text,
	W,
	xml,
} from "../src/docx-package.js";
import { executeFiller } from "../src/index.js";
import type { OfficePackage } from "../src/ooxml.js";
import { writeDocx } from "../src/pandoc.js";
import { renderPdfPages } from "../src/pdf.js";
import { runCommand } from "../src/process.js";
import { directory, enc, p, packageAt } from "./docx-edit-fixture.js";

// All document text, layout and raster data are fabricated here. No external assets.
function samplePng() {
	function crc(bytes: Buffer) {
		let result = 0xffffffff;
		for (const value of bytes) {
			result ^= value;
			for (let bit = 0; bit < 8; bit++) result = (result >>> 1) ^ (result & 1 ? 0xedb88320 : 0);
		}
		return (result ^ 0xffffffff) >>> 0;
	}
	function chunk(name: string, data: Buffer) {
		const kind = Buffer.from(name),
			length = Buffer.alloc(4),
			checksum = Buffer.alloc(4);
		length.writeUInt32BE(data.length);
		checksum.writeUInt32BE(crc(Buffer.concat([kind, data])));
		return Buffer.concat([length, kind, data, checksum]);
	}
	const header = Buffer.alloc(13);
	header.writeUInt32BE(8, 0);
	header.writeUInt32BE(16, 4);
	header[8] = 8;
	header[9] = 2;
	const pixels = Buffer.alloc(16 * (1 + 8 * 3));
	for (let y = 0; y < 16; y++)
		for (let x = 0; x < 8; x++) {
			const i = y * 25 + 1 + x * 3;
			pixels[i] = x * 30;
			pixels[i + 1] = y * 15;
			pixels[i + 2] = 180;
		}
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk("IHDR", header),
		chunk("IDAT", deflateSync(pixels)),
		chunk("IEND", Buffer.alloc(0)),
	]);
}

function syntheticTable() {
	const rows = Array.from(
		{ length: 4 },
		(_, index) =>
			`<w:tr><w:trPr><w:trHeight w:val="240"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(`Row ${String.fromCharCode(65 + index)}`)}</w:tc><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:shd w:fill="222222"/></w:tcPr>${p("")}</w:tc></w:tr>`,
	).join("");
	const borders = ["top", "left", "bottom", "right", "insideH", "insideV"]
		.map((name) => `<w:${name} w:val="single" w:sz="4"/>`)
		.join("");
	return `<w:tbl><w:tblPr><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>${rows}</w:tbl>`;
}

test("synthetic styled shell composes, retains package assets and renders to PDF", {
	timeout: 120000,
}, async (t) => {
	for (const executable of ["pandoc", "libreoffice", "pdftotext", "pdftocairo"]) {
		try {
			const result = await runCommand(executable, [
				executable === "pdftotext" || executable === "pdftocairo" ? "-v" : "--version",
			]);
			if (result.code !== 0) throw new Error("unavailable");
		} catch {
			t.skip("Optional local document-conversion executables unavailable");
			return;
		}
	}
	const dir = await directory();
	if (!process.env.PI_FILLER_KEEP_SYNTHETIC)
		t.after(() => rm(dir, { recursive: true, force: true }));
	else t.diagnostic(`Synthetic output directory: ${dir}`);
	const shellMarkdown = join(dir, "shell.md"),
		bodyMarkdown = join(dir, "body.md"),
		asset = join(dir, "asset.png");
	const shell = join(dir, "shell.docx"),
		donor = join(dir, "body.docx"),
		shaped = join(dir, "shaped.docx"),
		composed = join(dir, "composed.docx"),
		final = join(dir, "final.docx"),
		pdf = join(dir, "final.pdf");
	await writeFile(asset, samplePng());
	await writeFile(
		shellMarkdown,
		"# Sample shell\n\n![Sample icon](asset.png){width=0.5cm}\n\nREPLACE_BODY\n\nFabricated obsolete suffix.\n",
	);
	await writeFile(
		bodyMarkdown,
		"## Body section\n\nLead block\n\n![Sample drawing](asset.png){width=6cm}\n\nMiddle detail\n\nFinal detail\n",
	);
	const conversionOptions = {
		cwd: dir,
		env: {
			...process.env,
			HOME: dir,
			XDG_DATA_HOME: join(dir, "data"),
			XDG_CONFIG_HOME: join(dir, "config"),
		},
	};
	await writeDocx(shellMarkdown, shell, conversionOptions);
	await writeDocx(bodyMarkdown, donor, conversionOptions);
	const pkg: OfficePackage = await packageAt(shell);
	const root = parse(pkg["word/document.xml"], "document");
	const body = children(root, W, "body")[0];
	const anchor = elements(body.children, body.ns).find(
		(node) => is(node, W, "p") && text(node) === "REPLACE_BODY",
	);
	assert.ok(anchor);
	const panel = `<w:p xmlns:w="${W}"><w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="MockPanel" style="width:180pt;height:28pt" fillcolor="#DDDDDD"><v:textbox><w:txbxContent>${p("Fabricated panel")}</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p>`;
	body.children.splice(
		body.children.indexOf(anchor.entry),
		0,
		parse(enc(panel), "panel").entry,
		parse(enc(`<w:tbl xmlns:w="${W}">${syntheticTable().slice(7, -8)}</w:tbl>`), "table").entry,
	);
	const section = children(body, W, "sectPr")[0];
	assert.ok(section);
	const references = parse(
		enc(
			`<w:sectPr xmlns:w="${W}" xmlns:r="${R}"><w:headerReference w:type="default" r:id="rIdMockHeader"/><w:footerReference w:type="default" r:id="rIdMockFooter"/></w:sectPr>`,
		),
		"section",
	);
	section.children.unshift(...references.children);
	pkg["word/document.xml"] = xml(root);
	pkg["word/header1.xml"] = enc(`<w:hdr xmlns:w="${W}">${p("Sample header")}</w:hdr>`);
	pkg["word/footer1.xml"] = enc(
		`<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
	);
	const rels = parse(pkg["word/_rels/document.xml.rels"], "relationships");
	const types = parse(pkg["[Content_Types].xml"], "types");
	for (const kind of ["header", "footer"]) {
		const id = kind === "header" ? "rIdMockHeader" : "rIdMockFooter";
		rels.children.push(
			parse(
				enc(`<Relationship xmlns="${REL}" Id="${id}" Type="${R}/${kind}" Target="${kind}1.xml"/>`),
				"relationship",
			).entry,
		);
		types.children.push(
			parse(
				enc(
					`<Override xmlns="${CT}" PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`,
				),
				"content-type",
			).entry,
		);
	}
	pkg["word/_rels/document.xml.rels"] = xml(rels);
	pkg["[Content_Types].xml"] = xml(types);
	await writeFile(shell, zipSync(pkg));
	const shellHash = hash(await readFile(shell)),
		donorHash = hash(await readFile(donor));
	const donorAudit = await inspectDocxLayout(donor);
	await patchDocxLayout(donor, shaped, {
		paragraphs: [
			{
				from: 1,
				to: donorAudit.parts[0].paragraphCount,
				clearStyle: true,
				clearRunStyle: true,
				font: "Liberation Serif",
				size: 12,
				line: 300,
			},
			{ from: 1, alignment: "center", underline: true, pageBreakBefore: true },
			{ from: 1, to: donorAudit.parts[0].paragraphCount - 1, keepNext: true },
			{ from: donorAudit.parts[0].paragraphCount, keepNext: false },
		],
		imageBounds: { width: 4, height: 4 },
	});
	const report = await composeDocx(shell, composed, {
		body_path: shaped,
		replace_from: "REPLACE_BODY",
		body_from: "Body section",
		expected_sha256: shellHash,
	});
	assert.ok(report.importedRelationships > 0);
	await patchDocxLayout(composed, final, {
		cloneRows: [{ table: 1, row: 4, count: 2, content: "keep", labels: ["Row E", "Row F"] }],
	});
	const output = await packageAt(final);
	for (const part of Object.keys(pkg).filter((name) => /^word\/(header|footer|media)/.test(name)))
		assert.deepEqual(output[part], pkg[part]);
	assert.ok(new TextDecoder().decode(output["word/styles.xml"]).includes("\n"));
	const outputText = new TextDecoder().decode(output["word/document.xml"]);
	assert.match(outputText, /MockPanel/);
	assert.doesNotMatch(outputText, /Fabricated obsolete suffix|REPLACE_BODY/);
	const result = await executeFiller(
		{ format: "pdf", action: "export", path: final, output: pdf },
		dir,
	);
	const evidence = JSON.parse(result.text);
	assert.equal(evidence.pageCount, 2);
	assert.match(evidence.backend.version, /LibreOffice/);
	const extracted = await runCommand("pdftotext", ["-layout", pdf, "-"]);
	assert.equal(extracted.code, 0);
	const pages = extracted.stdout.toString("utf8").split("\f");
	assert.match(pages[0], /Row F/);
	assert.match(pages[1], /Final detail/);
	assert.match(pages[1], /Sample header/);
	const images = await renderPdfPages(pdf, join(dir, "rendered"), { firstPage: 1, lastPage: 2 });
	assert.equal(images.length, 2);
	assert.equal(hash(await readFile(shell)), shellHash);
	assert.equal(hash(await readFile(donor)), donorHash);
});
