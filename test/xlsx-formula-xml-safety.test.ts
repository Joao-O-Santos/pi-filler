// biome-ignore-all format: compact synthetic OOXML fixtures
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const workbook = "xl/workbook.xml";
const sheet = "xl/worksheets/sheet1.xml";
const enc = (value: string) => new TextEncoder().encode(value);
const dec = (value: Uint8Array) => new TextDecoder().decode(value);

function fixture(): Record<string, Uint8Array> {
	const files: Record<string, string> = {
		"[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
		"_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		[workbook]: `<workbook xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
		[sheet]: `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before after</t></is></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>`,
	};
	return Object.fromEntries(Object.entries(files).map(([name, value]) => [name, enc(value)]));
}

// Expat via minidom parses bytes independently and retains document-level PIs.
function inspect(xml: Uint8Array, expression: string): unknown {
	const script = `import json, sys\nfrom xml.dom import minidom\nd = minidom.parseString(sys.stdin.buffer.read())\nprint(json.dumps(${expression}))`;
	return JSON.parse(execFileSync("python3", ["-c", script], { input: xml, encoding: "utf8", stdio: ["pipe", "pipe", "pipe"] }));
}

async function setup(t: TestContext, parts: Record<string, Uint8Array>) {
	const dir = await mkdtemp(join(tmpdir(), "xlsx-xml-safety-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const path = join(dir, "source.xlsx");
	const output = join(dir, "output.xlsx");
	await writeFile(path, zipSync(parts));
	await writeFile(output, "previous output");
	return { dir, path, output };
}
async function run(dir: string, path: string, output: string) {
	return executeFiller({ format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell: "B1", formula: "=1+1" } } as never, dir);
}
function validateParts(parts: Record<string, Uint8Array>) {
	for (const [name, bytes] of Object.entries(parts)) {
		if (name.endsWith(".xml") || name.endsWith(".rels")) {
			assert.doesNotThrow(() => inspect(bytes, "d.documentElement.tagName"), `invalid XML in ${name}`);
		}
	}
}

for (const kind of ["invalid UTF-8", "NUL"] as const) {
	test(`${kind} in an unedited cell is refused without publishing invalid XML`, async (t) => {
		const parts = fixture();
		const marker = enc("before after");
		const index = Buffer.from(parts[sheet]).indexOf(marker);
		assert.notEqual(index, -1);
		parts[sheet][index + 3] = kind === "NUL" ? 0 : 0xff;
		assert.throws(() => inspect(parts[sheet], "d.documentElement.tagName"), "fixture must be invalid XML");
		const { dir, path, output } = await setup(t, parts);
		const source = await readFile(path);
		let rejected = false;
		try { await run(dir, path, output); } catch { rejected = true; }
		assert.deepEqual(await readFile(path), source);
		if (rejected) {
			assert.equal(await readFile(output, "utf8"), "previous output");
		} else {
			const result = unzipSync(await readFile(output));
			validateParts(result);
			assert.deepEqual(result[sheet], parts[sheet], "invalid input bytes must not be silently repaired or replaced");
		}
	});
}

test("top-level processing instruction is preserved or safely refused", async (t) => {
	const parts = fixture();
	parts[sheet] = enc(`<?audit keep="yes"?>${dec(parts[sheet])}`);
	const pi = "[(n.target, n.data) for n in d.childNodes if n.nodeType == n.PROCESSING_INSTRUCTION_NODE]";
	assert.deepEqual(inspect(parts[sheet], pi), [["audit", 'keep="yes"']]);
	const { dir, path, output } = await setup(t, parts);
	const source = await readFile(path);
	let rejected = false;
	try { await run(dir, path, output); } catch { rejected = true; }
	assert.deepEqual(await readFile(path), source);
	if (rejected) {
		assert.equal(await readFile(output, "utf8"), "previous output");
		return;
	}
	const result = unzipSync(await readFile(output));
	validateParts(result);
	assert.deepEqual(inspect(result[sheet], pi), inspect(parts[sheet], pi));
	assert.deepEqual(inspect(result[sheet], `[n.firstChild.data for n in d.getElementsByTagNameNS("${main}", "f") if n.firstChild]`), ["1+1"]);
});

test("ordinary LibreOffice aca=false formula is accepted and preserved", async (t) => {
	const parts = fixture();
	parts[sheet] = enc(dec(parts[sheet]).replace('<c r="A1" t="inlineStr"><is><t>before after</t></is></c>', '<c r="A1"><f aca="false">2+3</f><v>5</v></c>'));
	const { dir, path, output } = await setup(t, parts);
	await run(dir, path, output);
	const result = unzipSync(await readFile(output));
	validateParts(result);
	const formulas = `[(n.getAttribute("aca"), n.firstChild.data) for n in d.getElementsByTagNameNS("${main}", "f")]`;
	assert.deepEqual(inspect(result[sheet], formulas), [["false", "2+3"], ["", "1+1"]]);
});

test("harmless LibreOffice workbook extension is accepted and preserved", async (t) => {
	const parts = fixture();
	const ext = '<extLst><ext uri="{7626C862-2A13-11E5-B345-FEFF819CDC9F}" xmlns:loext="http://schemas.libreoffice.org/"><loext:extCalcPr stringRefSyntax="CalcA1ExcelA1"/></ext></extLst>';
	parts[workbook] = enc(dec(parts[workbook]).replace('</workbook>', `${ext}</workbook>`));
	const expression = '[(n.namespaceURI, n.localName, n.getAttribute("stringRefSyntax"), n.parentNode.getAttribute("uri")) for n in d.getElementsByTagNameNS("http://schemas.libreoffice.org/", "extCalcPr")]';
	assert.deepEqual(inspect(parts[workbook], expression), [["http://schemas.libreoffice.org/", "extCalcPr", "CalcA1ExcelA1", "{7626C862-2A13-11E5-B345-FEFF819CDC9F}"]]);
	const { dir, path, output } = await setup(t, parts);
	await run(dir, path, output);
	const result = unzipSync(await readFile(output));
	validateParts(result);
	assert.deepEqual(inspect(result[workbook], expression), inspect(parts[workbook], expression));
});
