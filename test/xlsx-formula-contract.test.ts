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

test("rejects merged target, nonordinary formulas, orphaned chain, signatures and invalid XML", async () => {
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

function withCalculationChain() {
	const parts = fixture();
	parts['xl/calcChain.xml'] = enc(`<calcChain ${ns}><c r="C3" i="1"/></calcChain>`);
	parts['xl/_rels/workbook.xml.rels'] = enc(dec(parts['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="chain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'));
	parts['[Content_Types].xml'] = enc(dec(parts['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
	return parts;
}

test("formula edits discard a conventional calculation chain without calculating caches", async () => {
	for (const mode of ['single', 'batch', 'sheets']) {
		const parts = withCalculationChain();
		const { dir, path, output } = await setup(parts);
		const original = await readFile(path);
		const extra = mode === 'batch' ? { xlsx_formula: undefined, xlsx_formulas: [{ cell: 'B2', formula }] } : mode === 'sheets' ? { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: [{ sheet: 'First', formulas: [{ cell: 'B2', formula }] }, { sheet: 'Second', formulas: [{ cell: 'A1', formula }] }] } : {};
		await writeFile(output, 'existing output');
		const dry = await executeFiller(request(path, output, { ...extra, dry_run: true }), dir);
		assert.equal(await readFile(output, 'utf8'), 'existing output');
		assert.ok((dry.details.changedParts as string[]).includes('xl/calcChain.xml'));
		const result = await executeFiller(request(path, output, extra), dir);
		privateCheck(result);
		assert.match(JSON.stringify(result.details.warnings), /stale cache/);
		const after = unzipSync(await readFile(output));
		assert.equal(after['xl/calcChain.xml'], undefined);
		assert.ok(!dec(after['xl/_rels/workbook.xml.rels']).includes('calcChain'));
		assert.ok(!dec(after['[Content_Types].xml']).includes('calcChain'));
		assert.equal(parser.parse(dec(after['xl/_rels/workbook.xml.rels'])).Relationships.Relationship.length, 3);
		assert.equal(cell(dec(after['xl/worksheets/sheet1.xml']), 'B2').v, undefined);
		assert.equal(cell(dec(after['xl/worksheets/sheet1.xml']), 'C3').v, '7');
		assert.equal(parser.parse(dec(after['xl/workbook.xml'])).workbook.calcPr['@_forceFullCalc'], '1');
		for (const [part, bytes] of Object.entries(parts)) if (!(result.details.changedParts as string[]).includes(part)) assert.deepEqual(after[part], bytes, part);
		assert.deepEqual(await readFile(path), original);
	}
});

test("ambiguous or unsupported chain packaging refuses transactionally", async () => {
	const mutations = [
		(p: Record<string, Uint8Array>) => { delete p['xl/calcChain.xml']; },
		(p: Record<string, Uint8Array>) => { p['xl/alias.xml'] = p['xl/calcChain.xml']; p['xl/_rels/workbook.xml.rels'] = enc(dec(p['xl/_rels/workbook.xml.rels']).replace('Target="calcChain.xml"', 'Target="alias.xml"')); },
		(p: Record<string, Uint8Array>) => { p['xl/_rels/workbook.xml.rels'] = enc(dec(p['xl/_rels/workbook.xml.rels']).replace('Id="chain"', 'Id="one"')); },
		(p: Record<string, Uint8Array>) => { p['xl/_rels/workbook.xml.rels'] = enc(dec(p['xl/_rels/workbook.xml.rels']).replace('Id="chain"', 'Id="chain" TargetMode="External"')); },
		(p: Record<string, Uint8Array>) => { p['[Content_Types].xml'] = enc(dec(p['[Content_Types].xml']).replace('PartName="/xl/calcChain.xml"', 'PartName="/xl/alias.xml"')); },
		(p: Record<string, Uint8Array>) => { p['xl/_rels/calcChain.xml.rels'] = enc('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'); },
		(p: Record<string, Uint8Array>) => { p['[Content_Types].xml'] = enc(dec(p['[Content_Types].xml']).replace('<Types ', '<Types xmlns:x="urn:foreign" ')); },
		(p: Record<string, Uint8Array>) => { p['xl/_rels/workbook.xml.rels'] = enc(dec(p['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="extra" Type="urn:unknown" Target="calcChain.xml"/></Relationships>')); },
		(p: Record<string, Uint8Array>) => { p['[Content_Types].xml'] = enc(dec(p['[Content_Types].xml']).replace('</Types>', '<!--preserve me--></Types>')); },
		(p: Record<string, Uint8Array>) => { p['[Content_Types].xml'] = enc(dec(p['[Content_Types].xml']).replace('</Types>', 'PRIVATE_MIXED_CONTENT</Types>')); },
		(p: Record<string, Uint8Array>) => { p['xl/calcChain.xml'] = enc('<broken>'); },
		(p: Record<string, Uint8Array>) => { p['_xmlsignatures/sig.xml'] = enc('<Signature/>'); },
	];
	for (const mutate of mutations) {
		const parts = withCalculationChain(); mutate(parts);
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

test("atomic formula batch creates cells, replaces formulas, and preserves nontarget data", async () => {
	const { dir, path, output } = await setup(fixture(`<c r="B2" s="1" t="str"><f>${oldSecret}</f><v>${oldSecret}</v></c>`));
	const original = await readFile(path);
	const xlsx_formulas = [{ cell: 'b2', formula }, { cell: 'D2', formula: '=B2+1' }, { cell: 'A8', formula: '=D2*2' }];
	const call = request(path, output, { xlsx_formula: undefined, xlsx_formulas });
	const result = await executeFiller(call, dir);
	privateCheck(result);
	assert.deepEqual(result.details.cells, ['B2', 'D2', 'A8']);
	assert.equal(result.details.formulaCount, 3);
	assert.equal(result.details.cell, undefined);
	const before = unzipSync(original);
	const after = unzipSync(await readFile(output));
	assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort());
	for (const name of Object.keys(before)) if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(name)) assert.deepEqual(after[name], before[name], name);
	const sheet = dec(after['xl/worksheets/sheet1.xml']);
	for (const item of xlsx_formulas) {
		const target = cell(sheet, item.cell.toUpperCase());
		assert.equal(target.f, item.formula.slice(1));
		assert.equal(target.v, undefined);
		assert.equal(target.is, undefined);
		assert.equal(target['@_t'], undefined);
	}
	assert.equal(cell(sheet, 'B2')['@_s'], '1');
	for (const address of ['A1', 'C3']) assert.deepEqual(cell(sheet, address), cell(dec(before['xl/worksheets/sheet1.xml']), address));
	const workbook = parser.parse(dec(after['xl/workbook.xml'])).workbook;
	assert.ok(!Array.isArray(workbook.calcPr));
	assert.equal(workbook.calcPr['@_forceFullCalc'], '1');
	assert.deepEqual(await readFile(path), original);
	const published = await readFile(output);
	const dry = await executeFiller(request(path, output, { xlsx_formula: undefined, xlsx_formulas, dry_run: true }), dir);
	assert.equal(dry.details.written, false);
	privateCheck(dry);
	assert.deepEqual(await readFile(output), published);
});

test("batch errors are private and all-or-nothing including a later merged target", async () => {
	const parts = fixture();
	parts['xl/worksheets/sheet1.xml'] = enc(dec(parts['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="D4:E4"/></mergeCells></worksheet>'));
	const { dir, path, output } = await setup(parts);
	const original = await readFile(path);
	const first = { cell: 'B2', formula };
	for (const extra of [
		{ xlsx_formulas: [] }, { xlsx_formulas: null }, { xlsx_formulas: {} },
		{ xlsx_formulas: Array.from({ length: 1001 }, () => first) },
		{ xlsx_formulas: [first, { cell: 'b2', formula: '=1' }] },
		{ xlsx_formulas: [first, { cell: 'D4', formula }] },
		{ xlsx_formulas: [first, { cell: 'A0', formula }] },
		{ xlsx_formulas: [first, { cell: 'A4', formula: '=' }] },
		{ xlsx_formulas: [first, { cell: 'A4', formula, sheet: 'Second' }] },
		{ xlsx_formulas: [first], xlsx_formula: first },
		{ xlsx_formulas: [first], range: 'B2' },
		{ xlsx_formulas: [first], xlsx_patches: { bold: true } },
		{ xlsx_formulas: [first], action: 'read', view: 'structure' },
	]) for (const dry_run of [false, true]) {
		await writeFile(output, 'existing output');
		await failure(() => executeFiller(request(path, output, { xlsx_formula: undefined, ...extra, dry_run }), dir));
		assert.equal(await readFile(output, 'utf8'), 'existing output');
		assert.deepEqual(await readFile(path), original);
	}
});

test("formula batch accepts its 1000-target bound", async () => {
	const { dir, path, output } = await setup();
	const xlsx_formulas = Array.from({ length: 1000 }, (_, i) => ({ cell: `D${i + 1}`, formula: '=1' }));
	const result = await executeFiller(request(path, output, { xlsx_formula: undefined, xlsx_formulas, dry_run: true }), dir);
	assert.equal(result.details.formulaCount, 1000);
	await assert.rejects(readFile(output), { code: 'ENOENT' });
});

test("multi-worksheet formulas publish once with private structural counts and preserve unrelated parts", async () => {
	const { dir, path, output } = await setup();
	const original = await readFile(path);
	const groups = [
		{ sheet: 'first', formulas: [{ cell: 'b2', formula }, { cell: 'D5', formula }] },
		{ sheet: 'SECOND', formulas: [{ cell: 'a1', formula }, { cell: 'D5', formula }] },
	];
	const call = { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: groups };
	const dry = await executeFiller(request(path, output, { ...call, dry_run: true }), dir);
	privateCheck(dry);
	assert.equal(dry.details.written, false);
	await assert.rejects(readFile(output), { code: 'ENOENT' });
	const result = await executeFiller(request(path, output, call), dir);
	privateCheck(result);
	assert.equal(result.details.written, true);
	assert.equal(result.details.formulaCount, 4);
	assert.equal(result.details.sheet, undefined);
	assert.deepEqual(result.details.sheets, [
		{ sheet: 'First', cells: ['B2', 'D5'], formulaCount: 2 },
		{ sheet: 'Second', cells: ['A1', 'D5'], formulaCount: 2 },
	]);
	assert.deepEqual(result.details.changedParts, ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
	const after = unzipSync(await readFile(output));
	const before = unzipSync(original);
	for (const part of Object.keys(before)) if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'].includes(part)) assert.deepEqual(after[part], before[part]);
	for (const [index, group] of groups.entries()) for (const item of group.formulas) {
		const target = cell(dec(after[`xl/worksheets/sheet${index + 1}.xml`]), item.cell.toUpperCase());
		assert.equal(target.f, formula.slice(1));
		assert.equal(target.v, undefined);
		assert.equal(target.is, undefined);
		assert.equal(target['@_t'], undefined);
	}
	assert.equal(cell(dec(after['xl/worksheets/sheet1.xml']), 'B2')['@_s'], '1');
	for (const address of ['A1', 'C3']) assert.deepEqual(cell(dec(after['xl/worksheets/sheet1.xml']), address), cell(dec(before['xl/worksheets/sheet1.xml']), address));
	const calc = parser.parse(dec(after['xl/workbook.xml'])).workbook.calcPr;
	assert.ok(!Array.isArray(calc));
	assert.equal(calc['@_forceFullCalc'], '1');
	assert.equal(calc['@_fullCalcOnLoad'], '1');
	assert.equal(calc['@_calcMode'], 'auto');
	const published = await readFile(output);
	await executeFiller(request(path, output, { ...call, dry_run: true }), dir);
	assert.deepEqual(await readFile(output), published);
	assert.deepEqual(await readFile(path), original);
});

test("multi-worksheet failures never write, including later merges, aliases, and aggregate bounds", async () => {
	const parts = fixture();
	parts['xl/worksheets/sheet2.xml'] = enc(dec(parts['xl/worksheets/sheet2.xml']).replace('</worksheet>', '<mergeCells><mergeCell ref="D4:E4"/></mergeCells></worksheet>'));
	const { dir, path, output } = await setup(parts);
	const original = await readFile(path);
	const first = { sheet: 'First', formulas: [{ cell: 'B2', formula }] };
	const second = { sheet: 'Second', formulas: [{ cell: 'A1', formula }] };
	for (const groups of [
		[], null, {}, [first, { ...second, sheet: 'missing' }],
		[first, { ...second, formulas: [{ cell: 'D4', formula }] }],
		[first, first], [first, { ...second, sheet: 'fIRST' }],
		[first, { ...second, formulas: [{ cell: 'a1', formula }, { cell: 'A1', formula }] }],
		[first, { ...second, extra: formula }],
		[first, { ...second, formulas: [{ cell: 'A1', formula: '=' }] }],
		[first, { ...second, formulas: Array.from({ length: 1000 }, (_, i) => ({ cell: `F${i + 1}`, formula })) }],
	]) for (const dry_run of [false, true]) {
		await writeFile(output, 'existing output');
		await failure(() => executeFiller(request(path, output, { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: groups, dry_run }), dir));
		assert.equal(await readFile(output, 'utf8'), 'existing output');
		assert.deepEqual(await readFile(path), original);
	}
	for (const extra of [{ sheet: 'First' }, { xlsx_formula: first.formulas[0] }, { xlsx_formulas: first.formulas }, { range: 'A1' }, { xlsx_patches: { bold: true } }, { action: 'read', view: 'structure' }]) {
		await failure(() => executeFiller(request(path, output, { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: [first, second], ...extra }), dir));
		assert.equal(await readFile(output, 'utf8'), 'existing output');
	}
	const bounded = [first, { ...second, formulas: Array.from({ length: 999 }, (_, i) => ({ cell: `F${i + 1}`, formula: '=1' })) }];
	const dry = await executeFiller(request(path, output, { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: bounded, dry_run: true }), dir);
	assert.equal(dry.details.formulaCount, 1000);
});

test("multi-worksheet rejects worksheet part aliases and ambiguous case-insensitive workbook names", async () => {
	for (const mode of ['part', 'name']) {
		const parts = fixture();
		if (mode === 'part') parts['xl/_rels/workbook.xml.rels'] = enc(dec(parts['xl/_rels/workbook.xml.rels']).replace('Target="worksheets/sheet2.xml"', 'Target="worksheets/sheet1.xml"'));
		else parts['xl/workbook.xml'] = enc(dec(parts['xl/workbook.xml']).replace('name="Second"', 'name="FIRST"'));
		const { dir, path, output } = await setup(parts);
		await failure(() => executeFiller(request(path, output, { sheet: undefined, xlsx_formula: undefined, xlsx_formula_sheets: [{ sheet: 'First', formulas: [{ cell: 'B2', formula }] }, { sheet: 'Second', formulas: [{ cell: 'A1', formula }] }] }), dir));
		await assert.rejects(readFile(output), { code: 'ENOENT' });
	}
});

test("model-facing schema and prompt describe formula patch and limitations", () => {
	let tool: { promptGuidelines?: string[]; promptSnippet?: string; description?: string } | undefined;
	extension({ registerTool(definition: typeof tool) { tool = definition; } } as never);
	const props = fillerSchema.properties as unknown as Record<string, Record<string, unknown>>;
	assert.ok(props.xlsx_formula, 'schema must expose xlsx_formula');
	assert.equal(props.xlsx_formula_sheets.minItems, 1);
	assert.equal(props.xlsx_formula_sheets.maxItems, 1000);
	assert.match(JSON.stringify(props.xlsx_formula_sheets), /1000 total|1–1000 total/);
	assert.match(JSON.stringify(props.xlsx_formula_sheets), /mutually exclusive/i);
	assert.equal(props.xlsx_formulas.minItems, 1);
	assert.equal(props.xlsx_formulas.maxItems, 1000);
	assert.match(JSON.stringify(props.xlsx_formulas), /mutually exclusive/i);
	assert.match(JSON.stringify(props.xlsx_formulas), /never formulas or values/i);
	const description = JSON.stringify(props.xlsx_formula);
	assert.match(description, /cell/i);
	assert.match(description, /formula/i);
	assert.match(description, /=/);
	const prompt = `${tool?.description} ${tool?.promptSnippet} ${tool?.promptGuidelines?.join(' ')}`;
	assert.match(prompt, /xlsx_formulas/);
	assert.match(prompt, /xlsx_formula_sheets/);
	assert.match(prompt, /calcChain\.xml calculation order caches are discarded/);
	assert.match(prompt, /does not calculate dependent caches/);
	assert.match(prompt, /multi-worksheet/);
	assert.match(prompt, /duplicate case-normalized/);
	assert.match(prompt, /every target.*atomic/);
	assert.match(prompt, /recalculat/i);
	assert.match(prompt, /(?:not|no|without).*evaluat|(?:not|no|without).*calculat/i);
});

test("model-facing contract advertises the LibreOffice recalculation request", () => {
	let tool: { description?: string } | undefined;
	extension({ registerTool(definition: typeof tool) { tool = definition; } } as never);
	const props = fillerSchema.properties as unknown as Record<string, Record<string, unknown>>;
	assert.equal(props.xlsx_recalculate.type, "boolean");
	assert.match(String(props.xlsx_recalculate.description), /LibreOffice/);
	assert.match(String(tool?.description), /xlsx_recalculate/);
});
