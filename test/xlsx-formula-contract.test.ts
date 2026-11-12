// biome-ignore-all format: synthetic OOXML fixtures are intentionally compact
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";
import extension, { executeFiller, fillerSchema } from "../src/index.js";

const oldSecret = "OLD_PRIVATE_1d1a76";
const formula = '=IF(A1<2,"FORMULA_PRIVATE_1d1a76",A1&"!")';
const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);
const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_", parseTagValue: false });

function fixture(target = `<c r="B2" s="1" t="str"><v>${oldSecret}</v></c>`, other = '<c r="C3"><f>A1+1</f><v>7</v></c>'): Record<string, Uint8Array> {
	return Object.fromEntries(Object.entries({
		"[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
		"_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		"xl/workbook.xml": `<workbook ${ns} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="one"/><sheet name="Second" sheetId="2" r:id="two"/></sheets><calcPr calcMode="manual"/></workbook>`,
		"xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="one" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="two" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
		"xl/worksheets/sheet1.xml": `<worksheet ${ns}><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2">${target}</row><row r="3">${other}</row></sheetData></worksheet>`,
		"xl/worksheets/sheet2.xml": `<worksheet ${ns}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>other sheet</t></is></c></row></sheetData></worksheet>`,
		"xl/styles.xml": `<styleSheet ${ns}><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
		"custom/data.bin": "untouched binary part",
	}).map(([key, value]) => [key, enc(value)]));
}

async function setup(parts = fixture()) {
	const dir = await mkdtemp(join(tmpdir(), "formula-contract-"));
	const path = join(dir, "input.xlsx");
	const output = join(dir, "output.xlsx");
	await writeFile(path, zipSync(parts));
	return { dir, path, output };
}
function request(path: string, output: string, extra: Record<string, unknown> = {}) {
	return { format: "xlsx", action: "patch", path, output, sheet: "First", xlsx_formula: { cell: "B2", formula }, ...extra } as never;
}
function privateCheck(value: unknown) {
	const text = value instanceof Error ? `${value.name}: ${value.message} ${JSON.stringify(value)}` : JSON.stringify(value);
	assert.ok(!text.includes(oldSecret));
	assert.ok(!text.includes("FORMULA_PRIVATE_1d1a76"));
}
async function failure(call: () => Promise<unknown>) {
	let error: unknown;
	try { await call(); } catch (caught) { error = caught; }
	assert.ok(error, "expected rejection");
	privateCheck(error);
}
function cell(sheet: string, address: string): Record<string, unknown> {
	assert.equal(XMLValidator.validate(sheet), true);
	const rows = parser.parse(sheet).worksheet.sheetData.row;
	for (const row of Array.isArray(rows) ? rows : [rows]) {
		const cells = Array.isArray(row.c) ? row.c : [row.c];
		for (const item of cells) if (item?.["@_r"] === address) return item;
	}
	assert.fail(`missing cell ${address}`);
}

test("formula creation preserves style and other package data and requests recalculation", async () => {
	const { dir, path, output } = await setup();
	const input = await readFile(path);
	const result = await executeFiller(request(path, output), dir);
	privateCheck(result);
	assert.deepEqual(await readFile(path), input);
	const before = unzipSync(input);
	const after = unzipSync(await readFile(output));
	for (const name of Object.keys(before)) if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(name)) assert.deepEqual(after[name], before[name], name);
	const sheet = dec(after['xl/worksheets/sheet1.xml']);
	const target = cell(sheet, 'B2');
	assert.equal(target.f, formula.slice(1));
	assert.equal(target['@_s'], '1');
	assert.equal(target['@_t'], undefined);
	assert.equal(target.v, undefined);
	assert.equal(target.is, undefined);
	assert.deepEqual(cell(sheet, 'C3'), cell(dec(before['xl/worksheets/sheet1.xml']), 'C3'));
	const workbook = dec(after['xl/workbook.xml']);
	assert.equal(XMLValidator.validate(workbook), true);
	const calc = parser.parse(workbook).workbook.calcPr;
	assert.ok(['1', 'true'].includes(calc['@_fullCalcOnLoad']));
	assert.notEqual(calc['@_calcMode'], 'manual');
});

test("replaces existing formula and cached result, or creates an absent cell on another sheet", async () => {
	for (const mode of ['replace', 'cross-sheet']) {
		const { dir, path, output } = await setup(fixture(`<c r="B2" s="1" t="e"><f>OLD_PRIVATE_1d1a76</f><v>#VALUE!</v></c>`));
		const extra = mode === 'cross-sheet' ? { sheet: 'Second', xlsx_formula: { cell: 'D4', formula } } : {};
		privateCheck(await executeFiller(request(path, output, extra), dir));
		const parts = unzipSync(await readFile(output));
		const target = cell(dec(parts[`xl/worksheets/${mode === 'replace' ? 'sheet1' : 'sheet2'}.xml`]), mode === 'replace' ? 'B2' : 'D4');
		assert.equal(target.f, formula.slice(1));
		assert.equal(target.v, undefined);
		assert.equal(target['@_t'], undefined);
		if (mode === 'replace') assert.equal(target['@_s'], '1');
	}
});

test("dry run validates but never overwrites output", async () => {
	const { dir, path, output } = await setup();
	await writeFile(output, 'existing output');
	privateCheck(await executeFiller(request(path, output, { dry_run: true }), dir));
	assert.equal(await readFile(output, 'utf8'), 'existing output');
});

test("invalid parameters reject privately without changing output or input", async () => {
	const { dir, path, output } = await setup();
	const original = await readFile(path);
	for (const extra of [
		{ xlsx_formula: { cell: 'A1:B2', formula } }, { xlsx_formula: { cell: '$A$1', formula } },
		{ xlsx_formula: { cell: 'XFE1', formula } }, { xlsx_formula: { cell: '', formula } },
		{ xlsx_formula: { cell: 'B2', formula: '' } }, { xlsx_formula: { cell: 'B2', formula: '=' } },
		{ xlsx_formula: { cell: 'B2', formula: 'SUM(A1)' } }, { xlsx_formula: { cell: 'B2' } },
		{ xlsx_formula: { cell: 'B2', formula, extra: true } },
		{ range: 'B2:B2', xlsx_patches: { row_height: 20 } },
		{ range: 'B2:B2' }, { xlsx_patches: { row_height: 20 } }, { sheet: 'missing' },
	]) {
		await writeFile(output, 'existing output');
		await failure(() => executeFiller(request(path, output, extra), dir));
		assert.equal(await readFile(output, 'utf8'), 'existing output');
		assert.deepEqual(await readFile(path), original);
	}
	await failure(() => executeFiller(request(path, path), dir));
	assert.deepEqual(await readFile(path), original);
});

test("rejects merged target, nonordinary formulas anywhere, calc chain, signatures and invalid XML", async () => {
	const cases: Record<string, Uint8Array>[] = [];
	const merged = fixture();
	merged['xl/worksheets/sheet1.xml'] = enc(dec(merged['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells></worksheet>'));
	cases.push(merged);
	for (const type of ['shared', 'array', 'dataTable', 'unexpected']) {
		const parts = fixture();
		parts['xl/worksheets/sheet2.xml'] = enc(`<worksheet ${ns}><sheetData><row r="1"><c r="A1"><f t="${type}" ref="A1:B2" si="0">${oldSecret}</f><v>1</v></c></row></sheetData></worksheet>`);
		cases.push(parts);
	}
	const chain = fixture(); chain['xl/calcChain.xml'] = enc(`<calcChain ${ns}><c r="A1" i="1"/></calcChain>`); cases.push(chain);
	const signed = fixture(); signed['_xmlsignatures/sig1.xml'] = enc('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>'); cases.push(signed);
	const malformed = fixture(); malformed['xl/worksheets/sheet2.xml'] = enc(`<worksheet>${oldSecret}<broken></worksheet>`); cases.push(malformed);
	for (const parts of cases) {
		const { dir, path, output } = await setup(parts);
		const original = await readFile(path);
		for (const dry_run of [false, true]) {
			await writeFile(output, 'existing output');
			await failure(() => executeFiller(request(path, output, { dry_run }), dir));
			assert.equal(await readFile(output, 'utf8'), 'existing output');
			assert.deepEqual(await readFile(path), original);
		}
	}
});

test("model-facing schema and prompt describe formula patch and limitations", () => {
	let tool: { promptGuidelines?: string[]; promptSnippet?: string; description?: string } | undefined;
	extension({ registerTool(definition: typeof tool) { tool = definition; } } as never);
	const props = fillerSchema.properties as unknown as Record<string, Record<string, unknown>>;
	assert.ok(props.xlsx_formula, 'schema must expose xlsx_formula');
	const description = JSON.stringify(props.xlsx_formula);
	assert.match(description, /cell/i);
	assert.match(description, /formula/i);
	assert.match(description, /=/);
	const prompt = `${tool?.description} ${tool?.promptSnippet} ${tool?.promptGuidelines?.join(' ')}`;
	assert.match(prompt, /xlsx_formula/);
	assert.match(prompt, /recalculat/i);
	assert.match(prompt, /(?:not|no|without).*evaluat|(?:not|no|without).*calculat/i);
});
