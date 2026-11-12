// biome-ignore-all format: compact synthetic OOXML fixtures
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";
import { fillXlsxFromCsv, inspectXlsxStructure, validateXlsxFormula } from "../src/xlsx.js";

const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const sheet = "xl/worksheets/sheet1.xml";
const workbook = "xl/workbook.xml";
const rels = "xl/_rels/workbook.xml.rels";
const enc = (value: string) => new TextEncoder().encode(value);
const dec = (value: Uint8Array) => new TextDecoder().decode(value);
function fixture(): Record<string, Uint8Array> {
	return Object.fromEntries(Object.entries({
		"[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
		"_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		[workbook]: `<workbook ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>`,
		[rels]: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
		[sheet]: `<worksheet ${ns}><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>`,
	}).map(([name, value]) => [name, enc(value)]));
}
async function setup(parts = fixture()) {
	const dir = await mkdtemp(join(tmpdir(), "formula-coverage-"));
	const path = join(dir, "input.xlsx");
	const output = join(dir, "output.xlsx");
	await writeFile(path, zipSync(parts));
	return { dir, path, output };
}
function request(path: string, output: string, extra: Record<string, unknown> = {}) {
	return { format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell: "B1", formula: "=A1+1" }, ...extra } as never;
}

test("formula request checks XML characters and UTF-16 length boundaries", () => {
	for (const formula of ["=\t1\n+2", `=${"x".repeat(8191)}`, "=😀"]) validateXlsxFormula({ cell: "A1", formula });
	for (const value of [null, [], {}, { cell: 1, formula: "=1" }, { cell: "A1", formula: 1 }, ...["= ", "=a\rb", "=\u0000", "=\ud800", `=${"x".repeat(8192)}`].map((formula) => ({ cell: "A1", formula }))]) {
		assert.throws(() => validateXlsxFormula(value), /Invalid XLSX formula request/);
	}
});

const invalidSheets: Array<[string, string, RegExp]> = [
	["missing sheetData", "", /explicit sheetData/],
	["unexpected sheetData child", "<sheetData><other/></sheetData>", /row structure/],
	["scalar row", "<sheetData><row>text</row></sheetData>", /row structure/],
	["missing row address", '<sheetData><row><c r="A1"/></row></sheetData>', /row structure/],
	["duplicate rows", '<sheetData><row r="1"/><row r="1"/></sheetData>', /row structure/],
	["unexpected row child", '<sheetData><row r="1"><other/></row></sheetData>', /cell structure/],
	["scalar cell", '<sheetData><row r="1"><c>text</c></row></sheetData>', /cell structure/],
	["cell on wrong row", '<sheetData><row r="1"><c r="A2"/></row></sheetData>', /cell structure/],
	["cells out of order", '<sheetData><row r="1"><c r="B1"/><c r="A1"/></row></sheetData>', /cell structure/],
	["cell metadata", '<sheetData><row r="1"><c r="A1" vm="1"/></row></sheetData>', /metadata/],
	["formula attribute", '<sheetData><row r="1"><c r="A1"><f aca="1">1</f></c></row></sheetData>', /ordinary formulas/],
	["repeated value", '<sheetData><row r="1"><c r="A1"><v>1</v><v>2</v></c></row></sheetData>', /Ambiguous XLSX XML/],
	["extension metadata", '<sheetData/><extLst/>', /metadata/],
];
for (const [name, content, error] of invalidSheets) {
	test(`rejects ${name} before replacing output`, async () => {
		const parts = fixture();
		parts[sheet] = enc(`<worksheet ${ns}>${content}</worksheet>`);
		const { dir, path, output } = await setup(parts);
		const original = await readFile(path);
		await writeFile(output, "preserve");
		await assert.rejects(executeFiller(request(path, output), dir), error);
		assert.equal(await readFile(output, "utf8"), "preserve");
		assert.deepEqual(await readFile(path), original);
	});
}

test("package-wide relationship, content type, and duplicate sheet guards reject transactionally", async () => {
	const cases: Array<[string, (p: Record<string, Uint8Array>) => void, RegExp]> = [
		["duplicate relationship", (p) => { p[rels] = enc(dec(p[rels]).replace('</Relationships>', '<Relationship Id="s1" TargetMode="External" Target="https://example.invalid"/></Relationships>')); }, /relationship structure/],
		["relationship metadata", (p) => { p[rels] = enc(dec(p[rels]).replace('</Relationships>', '<Relationship Id="other" Type="http://example/sheetMetadata" TargetMode="External" Target="https://example.invalid"/></Relationships>')); }, /metadata/],
		["content type signature", (p) => { p["[Content_Types].xml"] = enc(dec(p["[Content_Types].xml"]).replace('</Types>', '<Default Extension="sig" ContentType="application/digital-signature"/></Types>')); }, /signatures/],
		["duplicate sheet", (p) => { p[workbook] = enc(dec(p[workbook]).replace('</sheets>', '<sheet name="Data" sheetId="2" r:id="s1"/></sheets>')); }, /sheet structure/],
	];
	for (const [name, mutate, error] of cases) {
		const parts = fixture(); mutate(parts);
		const { dir, path, output } = await setup(parts);
		await writeFile(output, "preserve");
		await assert.rejects(executeFiller(request(path, output), dir), error, name);
		assert.equal(await readFile(output, "utf8"), "preserve", name);
	}
});

test("normal existing formulas and recalculation settings survive", async () => {
	const parts = fixture();
	parts[sheet] = enc(`<worksheet ${ns}><sheetData><row r="1"><c r="A1"><f t="normal">2+3</f><v>5</v></c></row></sheetData></worksheet>`);
	parts[workbook] = enc(dec(parts[workbook]).replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'));
	const { dir, path, output } = await setup(parts);
	await executeFiller(request(path, output), dir);
	const result = unzipSync(await readFile(output));
	assert.match(dec(result[workbook]), /calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/);
	assert.match(dec(result[sheet]), /<f t="normal">2\+3<\/f><v>5<\/v>/);
	assert.match(dec(result[sheet]), /<f>A1\+1<\/f>/);
});

test("structure inspection groups hidden rows and columns after a formula edit", async () => {
	const parts = fixture();
	parts[sheet] = enc(`<worksheet ${ns}><cols><col min="2" max="3" hidden="1"/></cols><sheetData><row r="1" hidden="1"><c r="A1"><v>2</v></c></row><row r="2" hidden="true"/><row r="4" hidden="1"/></sheetData></worksheet>`);
	const { dir, path, output } = await setup(parts);
	await executeFiller(request(path, output), dir);
	const structure = await inspectXlsxStructure(output);
	assert.deepEqual(structure.worksheets[0].hiddenRowRanges, ["1:2", "4"]);
	assert.deepEqual(structure.worksheets[0].hiddenColumnRanges, ["B:C"]);
});

test("formula target validation refuses invalid merge references", async () => {
	for (const merge of ['<mergeCell other="x"/>', '<mergeCell ref="C2:A1"/>']) {
		const parts = fixture();
		parts[sheet] = enc(dec(parts[sheet]).replace("</worksheet>", `<mergeCells>${merge}</mergeCells></worksheet>`));
		const { dir, path, output } = await setup(parts);
		await assert.rejects(executeFiller(request(path, output), dir), /merged range/);
		await assert.rejects(readFile(output), { code: "ENOENT" });
	}
});

test("CSV filling protects every formula and merged destination after formula editing", async () => {
	const parts = fixture();
	parts[sheet] = enc(`<worksheet ${ns}><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="B1"><f>2+2</f><v>4</v></c></row></sheetData><mergeCells><mergeCell ref="A2:B2"/></mergeCells></worksheet>`);
	const { dir, path, output } = await setup(parts);
	const edited = join(dir, "edited.xlsx");
	await executeFiller(request(path, edited, { xlsx_formula: { cell: "C3", formula: "=A1+B1" } }), dir);
	const csv = join(dir, "values.csv");
	await writeFile(csv, "one,two\n10,20\n");
	await writeFile(output, "preserve");
	await assert.rejects(fillXlsxFromCsv(edited, csv, output, { sheet: "Data", startCell: "A1" }), /intersects 2 formula cells/);
	await assert.rejects(fillXlsxFromCsv(edited, csv, output, { sheet: "Data", startCell: "A2" }), /intersects 1 merged range/);
	assert.equal(await readFile(output, "utf8"), "preserve");
});

test("formula dry run accepts nested absent output directories without creating them", async () => {
	const { dir, path } = await setup();
	const output = join(dir, "missing", "nested", "out.xlsx");
	await executeFiller(request(path, output, { dry_run: true }), dir);
	await assert.rejects(readFile(output), { code: "ENOENT" });
});

test("structure inspection rejects corrupt cell and hidden-column references", async () => {
	for (const [content, error] of [
		['<sheetData><row r="1"><c r="XFE1"><v>1</v></c></row></sheetData>', /invalid cell reference/],
		['<cols><col min="0" max="2" hidden="1"/></cols><sheetData/>', /invalid hidden column/],
	] as const) {
		const parts = fixture();
		parts[sheet] = enc(`<worksheet ${ns}>${content}</worksheet>`);
		const { path } = await setup(parts);
		await assert.rejects(inspectXlsxStructure(path), error);
	}
});

test("CSV mapping errors do not destroy a previously edited formula workbook", async () => {
	const { dir, path } = await setup();
	const edited = join(dir, "edited.xlsx");
	await executeFiller(request(path, edited), dir);
	const csv = join(dir, "values.csv");
	await writeFile(csv, "name,name\na,b\n");
	const output = join(dir, "result.xlsx");
	await writeFile(output, "preserve");
	for (const [columnMap, hasHeader, error] of [
		[{ name: "A" }, true, /headers must be unique/],
		[{ name: "A" }, false, /requires has_header/],
		[{ "3": "A" }, true, /ordinal is outside/],
		[{ "1": "A", name: "B" }, true, /all ordinals or all headers/],
	] as const) {
		await assert.rejects(fillXlsxFromCsv(edited, csv, output, { sheet: "Data", startCell: "A3", columnMap, hasHeader }), error);
		assert.equal(await readFile(output, "utf8"), "preserve");
	}
});

test("LibreOffice-produced workbook with unsupported namespaces is refused without output loss", async () => {
	const input = await readFile(new URL("./fixtures/libreoffice.xlsx", import.meta.url));
	const parts = unzipSync(input);
	assert.match(dec(parts[workbook]), /<extLst>/);
	const { dir, path, output } = await setup(parts);
	const original = await readFile(path);
	await writeFile(output, "preserve");
	await assert.rejects(executeFiller(request(path, output, { sheet: "libreoffice" }), dir), /Unsupported XLSX formula XML namespace/);
	assert.equal(await readFile(output, "utf8"), "preserve");
	assert.deepEqual(await readFile(path), original);
});
