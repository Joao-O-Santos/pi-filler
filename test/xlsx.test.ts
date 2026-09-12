// biome-ignore-all format: fixture-heavy regression tests preserve readable XML literals

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import test from "node:test";
import { promisify } from "node:util";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";
import {
	columnNameToNumber,
	columnNumberToName,
	fillXlsxFromCsv,
	inspectXlsxStructure,
	offsetCell,
	parseCellAddress,
	patchXlsx,
	rangeFromDimensions,
} from "../src/xlsx.js";

const SECRET = "DO_NOT_EXPOSE_31f0a4";
const executeFile = promisify(execFile);

function xml(value: string): Uint8Array<ArrayBuffer> {
	return Uint8Array.from(new TextEncoder().encode(`<?xml version="1.0"?>${value}`));
}

async function writeXlsxFixture(path: string, formulaInTarget = false): Promise<void> {
	const files = {
		"[Content_Types].xml": xml(
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
		),
		"_rels/.rels": xml(
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="root7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		),
		"xl/workbook.xml": xml(
			'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Results" sheetId="7" r:id="oddRelationship9"/><sheet name="Métadonnées utiles" sheetId="3" r:id="another2"/></sheets></workbook>',
		),
		"xl/_rels/workbook.xml.rels": xml(
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="another2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="style42" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="oddRelationship9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
		),
		"xl/worksheets/sheet1.xml": xml(
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>untouched</t></is></c></row></sheetData></worksheet>',
		),
		"xl/worksheets/sheet2.xml": xml(
			`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C10"/><cols><col min="1" max="3" hidden="1" style="1" width="9"/></cols><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>${SECRET}</t></is></c></row><row r="2"><c r="A2" s="1" t="inlineStr"><is><t>template</t></is></c>${formulaInTarget ? `<c r="B2"><f>${SECRET}+1</f><v>2</v></c>` : ""}</row><row r="10"><c r="C10"><f>${SECRET}+2</f><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells></worksheet>`,
		),
		"xl/worksheets/_rels/sheet2.xml.rels": xml(
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comment" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/><Relationship Id="drawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
		),
		"xl/comments1.xml": xml(
			`<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>${SECRET}</author></authors><commentList><comment ref="A1" authorId="0"><text><t>${SECRET}</t></text></comment></commentList></comments>`,
		),
		"xl/drawings/drawing1.xml": xml(
			'<drawing xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
		),
		"xl/styles.xml": xml(
			'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>',
		),
		"docProps/core.xml": xml(
			'<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><title>preserve metadata</title></coreProperties>',
		),
		"custom/preserved.bin": Uint8Array.from([9, 3, 7, 1]),
	};
	await writeFile(path, zipSync(files));
}

async function exists(path: string): Promise<boolean> {
	try {
		await stat(path);
		return true;
	} catch {
		return false;
	}
}

function packageContents(path: string): Promise<Record<string, Uint8Array>> {
	return readFile(path).then((bytes) => unzipSync(new Uint8Array(bytes)));
}

function packagePart(path: string, part: string): Promise<Uint8Array> {
	return packageContents(path).then((contents) => contents[part]);
}

function prefixedXml(bytes: Uint8Array, prefix: string): Uint8Array {
	const source = new TextDecoder().decode(bytes).replace(
		/<(\/?)([A-Za-z][A-Za-z0-9_.-]*)/g,
		`<$1${prefix}:$2`,
	);
	return Uint8Array.from(
		new TextEncoder().encode(
			source.replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', `xmlns:${prefix}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`),
		),
	);
}

function exposed(value: unknown): boolean {
	return JSON.stringify(value).includes(SECRET);
}

test("converts bounded XLSX cell addresses", () => {
	assert.deepEqual(parseCellAddress("A1"), { column: 1, row: 1 });
	assert.deepEqual(parseCellAddress("z1"), { column: 26, row: 1 });
	assert.deepEqual(parseCellAddress("AA20"), { column: 27, row: 20 });
	assert.deepEqual(parseCellAddress("XFD1048576"), { column: 16384, row: 1048576 });
	assert.equal(columnNameToNumber("AB"), 28);
	assert.throws(() => columnNameToNumber("ZZZ"));
	assert.equal(columnNumberToName(28), "AB");
	assert.equal(offsetCell("B2", 12, 25), "AA14");
	assert.equal(rangeFromDimensions("A2", 142, 8), "A2:H143");
	for (const invalid of ["", "A0", "1A", "$A$1", "XFE1", "A1048577", "A-1"]) {
		assert.throws(() => parseCellAddress(invalid));
	}
	assert.throws(() => columnNumberToName(0));
	assert.throws(() => offsetCell("A1", -1, 0));
	assert.throws(() => rangeFromDimensions("A1", 0, 1));
});

test("inspects XLSX structure through relationship IDs without exposing content", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-structure-"));
	const input = join(directory, "input.xlsx");
	await writeXlsxFixture(input);
	const structure = await inspectXlsxStructure(input);
	assert.deepEqual(structure.sheets, ["Results", "Métadonnées utiles"]);
	assert.equal(structure.activeSheet, "Results");
	assert.deepEqual(structure.worksheets[0], {
		sheet: "Results",
		usedRange: "A1:C10",
		rows: 10,
		columns: 3,
		formulaCells: 1,
		mergedRanges: 1,
		styledCells: 2,
	});
	assert.equal(exposed(structure), false);
	const result = await executeFiller(
		{ format: "xlsx", action: "read", view: "structure", path: input },
		directory,
	);
	assert.equal(exposed(result), false);
});

test("handles prefixed worksheets and styles while reporting the resolved part", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-prefix-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	const contents = await (async () => {
		const source = join(directory, "source.xlsx");
		await writeXlsxFixture(source);
		return packageContents(source);
	})();
	contents["xl/worksheets/custom.xml"] = prefixedXml(contents["xl/worksheets/sheet2.xml"], "x");
	delete contents["xl/worksheets/sheet2.xml"];
	contents["xl/styles.xml"] = prefixedXml(contents["xl/styles.xml"], "x");
	contents["xl/_rels/workbook.xml.rels"] = Uint8Array.from(
		new TextEncoder().encode(
			new TextDecoder().decode(contents["xl/_rels/workbook.xml.rels"]).replace(
				'Target="worksheets/sheet2.xml"',
				'Target="worksheets/custom.xml"',
			),
		),
	);
	await writeFile(input, zipSync(contents));
	await writeFile(csv, "heading\nvalue\n");
	const result = await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "D2",
	});
	assert.deepEqual(result.changedParts, ["xl/worksheets/custom.xml"]);
	const patched = await patchXlsx(output, join(directory, "patched.xlsx"), {
		sheet: "Results",
		range: "A1:A2",
		patches: { font: { bold: true }, fill: "D9EAF7" },
	});
	assert.deepEqual(patched.changedParts, ["xl/styles.xml", "xl/worksheets/custom.xml"]);
	const worksheet = new TextDecoder().decode(await packagePart(join(directory, "patched.xlsx"), "xl/worksheets/custom.xml"));
	assert.match(worksheet, /<x:c/);
	assert.doesNotMatch(worksheet, /<c[ >]/);
});

test("fills an XLSX template from complex CSV and preserves unrelated parts", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-write-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(
		csv,
		`code,note,flag\n00123,"comma, quote ""kept"" and\nnewline",true\n,"café ${SECRET}",false\n`,
	);
	const before = await packageContents(input);
	const result = await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "A2",
	});
	assert.deepEqual(
		{ range: result.writtenRange, rows: result.rows, columns: result.columns, written: result.written },
		{ range: "A2:C3", rows: 2, columns: 3, written: true },
	);
	const after = await packageContents(output);
	for (const [part, bytes] of Object.entries(before)) {
		if (part !== "xl/worksheets/sheet2.xml") assert.deepEqual(after[part], bytes, part);
	}
	const worksheet = new TextDecoder().decode(await packagePart(output, "xl/worksheets/sheet2.xml"));
	assert.match(worksheet, /r="A2" s="1" t="inlineStr"/);
	assert.match(worksheet, />00123</);
	assert.match(worksheet, /comma, quote &quot;kept&quot; and\nnewline/);
	assert.match(worksheet, new RegExp(`café ${SECRET}`));
	assert.match(worksheet, new RegExp(`<f>${SECRET}\\+2</f>`));
	assert.equal(exposed(result), false);
});

test("supports no-header CSV and conservative automatic primitives", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-auto-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(csv, "00123,42,true,=A1,1.5,\n");
	await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "D2",
		hasHeader: false,
		valueMode: "auto",
	});
	const worksheet = new TextDecoder().decode(await packagePart(output, "xl/worksheets/sheet2.xml"));
	assert.match(worksheet, /r="D2" t="inlineStr"[^>]*>.*00123/);
	assert.match(worksheet, /r="E2"><v>42<\/v>/);
	assert.match(worksheet, /r="F2" t="b"><v>1<\/v>/);
	assert.match(worksheet, /r="G2" t="inlineStr"[^>]*>.*=A1/);
	assert.match(worksheet, /r="H2"><v>1.5<\/v>/);
	assert.match(worksheet, /r="I2"\/>|r="I2"><\/c>/);
});

test("handles a large bounded CSV without returning its values", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-large-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	const records = Array.from({ length: 5_000 }, (_, row) =>
		Array.from({ length: 8 }, (_, column) => `private-${row}-${column}`).join(","),
	);
	await writeFile(csv, records.join("\n"));
	const result = await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "D1",
		hasHeader: false,
		dryRun: true,
	});
	assert.deepEqual(
		{ rows: result.rows, columns: result.columns, range: result.targetRange },
		{ rows: 5_000, columns: 8, range: "D1:K5000" },
	);
	assert.equal(JSON.stringify(result).includes("private-"), false);
});

test("refuses formula replacement transactionally without exposing formulas or CSV", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-formula-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input, true);
	await writeFile(csv, `a,b\n${SECRET},replacement\n`);
	await writeFile(output, "prior output");
	let failure: unknown;
	try {
		await fillXlsxFromCsv(input, csv, output, { sheet: "Results", startCell: "A2" });
	} catch (error) {
		failure = error;
	}
	assert.match(String(failure), /intersects 1 formula cell/);
	assert.equal(String(failure).includes(SECRET), false);
	assert.equal(await readFile(output, "utf8"), "prior output");
});

test("protects formula-owned array ranges beyond their anchor cell", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-array-formula-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input, true);
	const contents = await packageContents(input);
	const sheet = new TextDecoder().decode(contents["xl/worksheets/sheet2.xml"]);
	contents["xl/worksheets/sheet2.xml"] = Uint8Array.from(
		new TextEncoder().encode(sheet.replace(`<f>${SECRET}+1</f>`, `<f ref="B2:C3">${SECRET}+1</f>`)),
	);
	await writeFile(input, zipSync(contents));
	await writeFile(csv, "header\nreplacement\n");
	await assert.rejects(
		() => fillXlsxFromCsv(input, csv, output, { sheet: "Results", startCell: "C3" }),
		/intersects 1 formula cell/,
	);
	assert.equal(await exists(output), false);
});

test("CSV dry runs validate the same mutation as real writes", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-dry-parity-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(csv, "heading,value\n00123,42\n");
	const dry = await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "D2",
		dryRun: true,
	});
	const real = await fillXlsxFromCsv(input, csv, output, {
		sheet: "Results",
		startCell: "D2",
	});
	assert.equal(dry.written, false);
	assert.equal(real.written, true);
	assert.equal(dry.targetRange, real.targetRange);
	assert.equal(dry.changedParts[0], real.changedParts[0]);
});

test("preserves extLst ordering and XML Schema row style inheritance", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-order-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const contents = await (async () => {
		const source = join(directory, "source.xlsx");
		await writeXlsxFixture(source);
		return packageContents(source);
	})();
	let sheet = new TextDecoder().decode(contents["xl/worksheets/sheet2.xml"]);
	sheet = sheet.replace('<row r="2">', '<row r="2" s="1" customFormat="true">');
	sheet = sheet.replace('<c r="A2" s="1"', '<c r="A2"');
	sheet = sheet.replace(
		/(<c r="A2"[^>]*>[\s\S]*?<\/is>)(<\/c>)/,
		"$1<extLst/>$2",
	);
	contents["xl/worksheets/sheet2.xml"] = Uint8Array.from(new TextEncoder().encode(sheet));
	await writeFile(input, zipSync(contents));
	await patchXlsx(input, output, {
		sheet: "Results",
		range: "A2:A2",
		patches: { font: { italic: true } },
	});
	const worksheet = new TextDecoder().decode(await packagePart(output, "xl/worksheets/sheet2.xml"));
	assert.match(worksheet, /<is>[\s\S]*<extLst><\/extLst><\/c>/);
	const styles = new TextDecoder().decode(await packagePart(output, "xl/styles.xml"));
	assert.match(styles, /<font>[\s\S]*<b><\/b>[\s\S]*<i><\/i>[\s\S]*<\/font>/);
});

test("dry-run XLSX writes expose metadata only and do not write", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-write-dry-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(csv, `head\n${SECRET}\n`);
	const result = await executeFiller(
		{
			format: "xlsx",
			action: "write",
			path: input,
			output,
			source_csv: csv,
			sheet: "Results",
			start_cell: "A2",
			dry_run: true,
		},
		directory,
	);
	assert.equal(exposed(result), false);
	assert.equal(await exists(output), false);
});

test("patches XLSX styles and dimensions while preserving cell contents", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-patch-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	await writeXlsxFixture(input, true);
	const result = await patchXlsx(input, output, {
		sheet: "Results",
		range: "A1:B2",
		patches: {
			font: { bold: true, italic: true, size: 14 },
			fill: "D9EAF7",
			alignment: { horizontal: "center", vertical: "bottom", wrap_text: true },
			number_format: "0.00",
			border: { style: "thin" },
			column_width: 18,
			row_height: 24,
		},
	});
	assert.deepEqual(
		{
			styled: result.styledCells,
			columns: result.columnsChanged,
			rows: result.rowsChanged,
			parts: result.changedParts,
		},
		{
			styled: 4,
			columns: 2,
			rows: 2,
			parts: ["xl/styles.xml", "xl/worksheets/sheet2.xml"],
		},
	);
	assert.equal(exposed(result), false);
	const worksheet = new TextDecoder().decode(await packagePart(output, "xl/worksheets/sheet2.xml"));
	const styles = new TextDecoder().decode(await packagePart(output, "xl/styles.xml"));
	assert.match(worksheet, new RegExp(SECRET));
	assert.match(worksheet, /width="18"[^>]*customWidth="1"/);
	assert.match(worksheet, /min="1" max="1"[^>]*hidden="1"[^>]*style="1"/);
	assert.match(worksheet, /r="1" ht="24" customHeight="1"/);
	assert.match(styles, /rgb="FFD9EAF7"/);
	assert.match(styles, /horizontal="center" vertical="bottom" wrapText="1"/);
	assert.match(styles, /formatCode="0.00"/);
	assert.match(styles, /style="thin"/);
	assert.equal((await inspectXlsxStructure(output)).worksheets[0].formulaCells, 2);
	const before = await packageContents(input);
	const after = await packageContents(output);
	for (const [part, bytes] of Object.entries(before)) {
		if (!["xl/styles.xml", "xl/worksheets/sheet2.xml"].includes(part)) {
			assert.deepEqual(after[part], bytes, part);
		}
	}
});

test("column-only XLSX patches do not create rows or cells", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-columns-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	await writeXlsxFixture(input);
	await patchXlsx(input, output, {
		sheet: "Results",
		range: "A100:B200",
		patches: { column_width: 22 },
	});
	assert.equal((await inspectXlsxStructure(output)).worksheets[0].usedRange, "A1:C10");
});

test("reuses equivalent XLSX styles", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-style-reuse-"));
	const input = join(directory, "input.xlsx");
	const first = join(directory, "first.xlsx");
	const second = join(directory, "second.xlsx");
	await writeXlsxFixture(input);
	const patch = {
		sheet: "Results",
		range: "B1:B2",
		patches: { font: { bold: true } },
	};
	await patchXlsx(input, first, patch);
	await patchXlsx(first, second, patch);
	for (const path of [first, second]) {
		const styles = new TextDecoder().decode(await packagePart(path, "xl/styles.xml"));
		assert.match(styles, /<fonts count="2">/);
		assert.match(styles, /<cellXfs count="2">/);
	}
});

test("dry-run XLSX patches do not expose content or write output", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-patch-dry-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	await writeXlsxFixture(input);
	const result = await executeFiller(
		{
			format: "xlsx",
			action: "patch",
			path: input,
			output,
			sheet: "Results",
			range: "A1:A1",
			xlsx_patches: { font: { bold: true } },
			dry_run: true,
		},
		directory,
	);
	assert.equal(exposed(result), false);
	assert.equal(await exists(output), false);
});

test("mutates a LibreOffice-produced XLSX fixture and optionally reopens it", async (context) => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-libreoffice-"));
	const input = resolve("test/fixtures/libreoffice.xlsx");
	const filled = join(directory, "filled.xlsx");
	const patched = join(directory, "patched.xlsx");
	const csv = join(directory, "source.csv");
	const structure = await inspectXlsxStructure(input);
	assert.equal(structure.worksheets[0].usedRange, "A1:B2");
	await writeFile(csv, `heading,value\n${SECRET},7\n`);
	const writeResult = await fillXlsxFromCsv(input, csv, filled, {
		sheet: structure.sheets[0],
		startCell: "A2",
	});
	assert.equal(exposed(writeResult), false);
	const patchResult = await patchXlsx(filled, patched, {
		sheet: structure.sheets[0],
		range: "A1:B2",
		patches: { font: { bold: true }, column_width: 16 },
	});
	assert.equal(exposed(patchResult), false);
	assert.equal((await inspectXlsxStructure(patched)).worksheets[0].usedRange, "A1:B2");

	const profile = join(directory, "libreoffice-profile");
	try {
		await executeFile(
			"libreoffice",
			[
				`-env:UserInstallation=file://${profile}`,
				"--headless",
				"--convert-to",
				"csv",
				"--outdir",
				directory,
				patched,
			],
			{ timeout: 30_000 },
		);
		assert.equal(await exists(join(directory, "patched.csv")), true);
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		context.diagnostic("LibreOffice unavailable; package-level integration checks passed");
	}
});

test("keeps XLSX failures private and output-safe", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-private-errors-"));
	const input = join(directory, "input.xlsx");
	const output = join(directory, "output.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(csv, `header\n"${SECRET}`);
	await writeFile(output, "prior output");
	let malformedCsv: unknown;
	try {
		await fillXlsxFromCsv(input, csv, output, { sheet: "Results", startCell: "A2" });
	} catch (error) {
		malformedCsv = error;
	}
	assert.match(String(malformedCsv), /malformed CSV/);
	assert.equal(String(malformedCsv).includes(SECRET), false);
	assert.equal(await readFile(output, "utf8"), "prior output");

	const malformed = join(directory, "malformed.xlsx");
	const contents = await packageContents(input);
	contents["xl/worksheets/sheet2.xml"] = xml(`<worksheet><broken>${SECRET}</worksheet>`);
	await writeFile(malformed, zipSync(contents));
	let malformedXml: unknown;
	try {
		await inspectXlsxStructure(malformed);
	} catch (error) {
		malformedXml = error;
	}
	assert.match(String(malformedXml), /Invalid XML/);
	assert.equal(String(malformedXml).includes(SECRET), false);

	const controller = new AbortController();
	controller.abort();
	await assert.rejects(
		() =>
			patchXlsx(
				input,
				join(directory, "aborted.xlsx"),
				{ sheet: "Results", range: "A1:A1", patches: { row_height: 20 } },
				false,
				controller.signal,
			),
		/operation aborted/,
	);
	assert.equal(await exists(join(directory, "aborted.xlsx")), false);

	const outputDirectory = await mkdtemp(join(directory, "existing-output-"));
	await assert.rejects(() =>
		patchXlsx(input, outputDirectory, {
			sheet: "Results",
			range: "A1:A1",
			patches: { row_height: 20 },
		}),
	);
	assert.equal((await stat(outputDirectory)).isDirectory(), true);

	const noStyles = join(directory, "no-styles.xlsx");
	const withoutStyles = await packageContents(input);
	delete withoutStyles["xl/styles.xml"];
	withoutStyles["xl/_rels/workbook.xml.rels"] = xml(
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="another2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="oddRelationship9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
	);
	await writeFile(noStyles, zipSync(withoutStyles));
	await assert.rejects(
		() =>
			patchXlsx(
				noStyles,
				join(directory, "no-styles-output.xlsx"),
				{ sheet: "Results", range: "A1:A1", patches: { font: { bold: true } } },
				true,
			),
		/requires xl\/styles.xml/,
	);
});

test("rejects unsupported XLSX calls, missing sheets, bad packages, and same paths", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-xlsx-errors-"));
	const input = join(directory, "input.xlsx");
	const csv = join(directory, "source.csv");
	await writeXlsxFixture(input);
	await writeFile(csv, "h\nv\n");
	await assert.rejects(
		() => executeFiller({ format: "xlsx", action: "read", view: "text", path: input }, directory),
		/requires view=structure/,
	);
	await assert.rejects(
		() => executeFiller({ format: "xlsx", action: "search", path: input }, directory),
		/search is not supported/,
	);
	await assert.rejects(
		() =>
			executeFiller(
				{
					format: "xlsx",
					action: "write",
					path: input,
					output: join(directory, "out.xlsx"),
					source_csv: csv,
					sheet: "missing",
					start_cell: "A1",
					query: "irrelevant",
				},
				directory,
			),
		/does not accept query/,
	);
	await assert.rejects(
		() => fillXlsxFromCsv(input, csv, input, { sheet: "Results", startCell: "A1" }),
		/output must differ/,
	);
	await assert.rejects(
		() => fillXlsxFromCsv(input, csv, join(directory, "out.xlsx"), { sheet: "missing", startCell: "A1" }),
		/does not exist/,
	);
	await writeFile(join(directory, "bad.xlsx"), "not zip");
	await assert.rejects(() => inspectXlsxStructure(join(directory, "bad.xlsx")), /Invalid XLSX ZIP/);
});
