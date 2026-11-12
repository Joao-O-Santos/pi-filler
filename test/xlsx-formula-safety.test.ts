// biome-ignore-all format: compact synthetic OOXML fixtures
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const enc = (s: string) => new TextEncoder().encode(s);
const dec = (s: Uint8Array) => new TextDecoder().decode(s);
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_", parseTagValue: false, trimValues: false });
const secret = "PRIVATE_雪_4937";
const formula = `=  IF(A1<2,"${secret} & < > ' \"\"",A1&"é")  `;

function packageParts(): Record<string, Uint8Array> {
	const files: Record<string, string> = {
		"[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
		"_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		"xl/workbook.xml": `<workbook ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/><sheet name="Elsewhere" sheetId="2" r:id="s2"/></sheets><calcPr calcId="123" calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0"/></workbook>`,
		"xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="s2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
		"xl/worksheets/sheet1.xml": `<worksheet ${ns}><sheetData><row r="1"><c r="A1"><v>5</v></c><c r="C1"><v>8</v></c></row><row r="3"><c r="B3" s="2" t="inlineStr"><is><t>${secret}</t></is></c><c r="D3"><v>9</v></c></row><row r="5"><c r="A5"><v>6</v></c></row></sheetData></worksheet>`,
		"xl/worksheets/sheet2.xml": `<worksheet ${ns}><sheetData><row r="1"><c r="A1"><v>4</v></c></row></sheetData></worksheet>`,
	};
	return Object.fromEntries(Object.entries(files).map(([name, value]) => [name, enc(value)]));
}
async function setup(parts = packageParts()) {
	const dir = await mkdtemp(join(tmpdir(), "xlsx-safety-"));
	const path = join(dir, "source.xlsx");
	const output = join(dir, "result.xlsx");
	await writeFile(path, zipSync(parts));
	return { dir, path, output };
}
function request(path: string, output: string, cell = "B3", extra: Record<string, unknown> = {}) {
	return { format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell, formula }, ...extra } as never;
}
function noSecrets(value: unknown) {
	const text = value instanceof Error ? `${value.stack} ${JSON.stringify(value)}` : JSON.stringify(value);
	assert.ok(!text.includes(secret), "sensitive cell/formula text appeared in result or error");
	assert.ok(!text.includes(formula), "formula appeared in result or error");
}
function parsed(bytes: Uint8Array) {
	const xml = dec(bytes);
	assert.equal(XMLValidator.validate(xml), true);
	return parser.parse(xml);
}
function list<T>(value: T | T[]): T[] { return Array.isArray(value) ? value : [value]; }

test("literal replacement round-trips XML-sensitive Unicode and whitespace without exposing it", async () => {
	const { dir, path, output } = await setup();
	const before = await readFile(path);
	const result = await executeFiller(request(path, output), dir);
	noSecrets(result);
	assert.deepEqual(await readFile(path), before);
	const parts = unzipSync(await readFile(output));
	const rows = list<any>(parsed(parts["xl/worksheets/sheet1.xml"]).worksheet.sheetData.row);
	const target = list<any>(rows[1].c).find((c) => c["@_r"] === "B3");
	assert.equal(typeof target.f === "object" ? target.f["#text"] : target.f, formula.slice(1));
	assert.equal(target["@_s"], "2");
	assert.equal(target["@_t"], undefined);
	assert.equal(target.is, undefined);
	assert.equal(target.v, undefined);
	const calc = parsed(parts["xl/workbook.xml"]).workbook.calcPr;
	assert.notEqual(calc["@_calcMode"], "manual");
	assert.ok(["1", "true"].includes(calc["@_fullCalcOnLoad"]));
	assert.ok(["1", "true"].includes(calc["@_forceFullCalc"]));
});

test("new cells and rows remain in numeric worksheet order", async () => {
	for (const [address, expectedRows, rowNumber, expectedCells] of [
		["B1", ["1", "3", "5"], "1", ["A1", "B1", "C1"]],
		["C2", ["1", "2", "3", "5"], "2", ["C2"]],
		["C3", ["1", "3", "5"], "3", ["B3", "C3", "D3"]],
	] as const) {
		const { dir, path, output } = await setup();
		noSecrets(await executeFiller(request(path, output, address), dir));
		const rows = list<any>(parsed(unzipSync(await readFile(output))["xl/worksheets/sheet1.xml"]).worksheet.sheetData.row);
		assert.deepEqual(rows.map((r) => r["@_r"]), expectedRows);
		assert.deepEqual(list<any>(rows.find((r) => r["@_r"] === rowNumber).c).map((c) => c["@_r"]), expectedCells);
	}
});

test("dry run preserves existing output and returns no formula", async () => {
	const { dir, path, output } = await setup();
	const original = await readFile(path);
	await writeFile(output, "do not replace");
	noSecrets(await executeFiller(request(path, output, "B3", { dry_run: true }), dir));
	assert.equal(await readFile(output, "utf8"), "do not replace");
	assert.deepEqual(await readFile(path), original);
});

const rejected: Array<[string, (parts: Record<string, Uint8Array>) => void]> = [
	...["shared", "array", "dataTable"].map((type): [string, (parts: Record<string, Uint8Array>) => void] => [type, (parts) => {
		parts["xl/worksheets/sheet2.xml"] = enc(`<worksheet ${ns}><sheetData><row r="1"><c r="A1"><f t="${type}" ref="A1:B2" si="0">${secret}</f><v>1</v></c></row></sheetData></worksheet>`);
	}]),
	["cell metadata", (parts) => { parts["xl/metadata.xml"] = enc(`<metadata ${ns}><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000"/></metadataTypes></metadata>`); }],
	["calculation chain", (parts) => { parts["xl/calcChain.xml"] = enc(`<calcChain ${ns}><c r="A1" i="1"/></calcChain>`); }],
	["signature", (parts) => { parts["_xmlsignatures/sig1.xml"] = enc('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>'); }],
	["malformed other worksheet", (parts) => { parts["xl/worksheets/sheet2.xml"] = enc(`<worksheet>${secret}<bad></worksheet>`); }],
];
for (const [name, mutate] of rejected) {
	test(`refuses ${name} privately and transactionally, including dry run`, async () => {
		const parts = packageParts();
		mutate(parts);
		const { dir, path, output } = await setup(parts);
		const original = await readFile(path);
		for (const dry_run of [false, true]) {
			await writeFile(output, "previous output");
			let error: unknown;
			try { await executeFiller(request(path, output, "B3", { dry_run }), dir); } catch (caught) { error = caught; }
			assert.ok(error, "expected rejection");
			noSecrets(error);
			assert.equal(await readFile(output, "utf8"), "previous output");
			assert.deepEqual(await readFile(path), original);
		}
	});
}

test("same source and output is rejected without modifying source", async () => {
	const { dir, path } = await setup();
	const original = await readFile(path);
	let error: unknown;
	try { await executeFiller(request(path, path), dir); } catch (caught) { error = caught; }
	assert.ok(error);
	noSecrets(error);
	assert.deepEqual(await readFile(path), original);
});
