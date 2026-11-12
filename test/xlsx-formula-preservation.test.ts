// biome-ignore-all format: compact synthetic OOXML fixtures
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
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

async function setup(t: TestContext, parts = fixture()) {
	const dir = await mkdtemp(join(tmpdir(), "xlsx-preservation-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const path = join(dir, "source.xlsx");
	const output = join(dir, "output.xlsx");
	await writeFile(path, zipSync(parts));
	return { dir, path, output };
}
function request(path: string, output: string, formula = "=1+1") {
	return { format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell: "B1", formula } } as never;
}

// ElementTree is independent of the production XML parser and resolves expanded names.
function inspect(xml: Uint8Array, code: string): unknown {
	return JSON.parse(execFileSync("python3", ["-c", `import json, sys, xml.etree.ElementTree as E\nr=E.fromstring(sys.stdin.buffer.read())\nM=${JSON.stringify(`{${main}}`)}\n${code}`], { input: xml, encoding: "utf8" }));
}
async function editOrRefuse(dir: string, path: string, output: string, reason: RegExp, formula = "=1+1") {
	const original = await readFile(path);
	let rejected = false;
	try { await executeFiller(request(path, output, formula), dir); } catch (error) {
		rejected = true;
		assert.match(String(error), reason, "refusal must correspond to the tested condition");
	}
	assert.deepEqual(await readFile(path), original, "source must remain unchanged");
	if (rejected) {
		await assert.rejects(() => readFile(output), { code: "ENOENT" });
		return undefined;
	}
	const parts = unzipSync(await readFile(output));
	assert.equal(inspect(parts[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'B1\']/"+M+"f").text))'), formula.slice(1));
	return parts;
}

test("ordinary fixture accepts a formula edit", async (t) => {
	const { dir, path, output } = await setup(t);
	const original = await readFile(path);
	await executeFiller(request(path, output), dir);
	assert.deepEqual(await readFile(path), original);
	const parts = unzipSync(await readFile(output));
	assert.equal(inspect(parts[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'B1\']/"+M+"f").text))'), "1+1");
	assert.equal(inspect(parts[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'A1\']/"+M+"is/"+M+"t").text))'), "before after");
});

test("numeric CR references in unrelated cells retain their XML character value", async (t) => {
	const input = fixture();
	input[sheet] = enc(dec(input[sheet]).replace("before after", "before&#13;after"));
	assert.equal(inspect(input[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'A1\']/"+M+"is/"+M+"t").text))'), "before\rafter");
	const { dir, path, output } = await setup(t, input);
	const parts = await editOrRefuse(dir, path, output, /XML text representation/);
	if (!parts) return;
	const text = inspect(parts[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'A1\']/"+M+"is/"+M+"t").text))');
	assert.equal(text, "before\rafter");
});

test("literal CR in unrelated XML text retains its normalized character value or is refused", async (t) => {
	const input = fixture();
	input[sheet] = enc(dec(input[sheet]).replace("before after", "before\rafter"));
	const code = 'print(json.dumps(r.find(".//"+M+"c[@r=\'A1\']/"+M+"is/"+M+"t").text))';
	assert.equal(inspect(input[sheet], code), "before\nafter");
	const { dir, path, output } = await setup(t, input);
	const parts = await editOrRefuse(dir, path, output, /XML text representation/);
	if (parts) assert.equal(inspect(parts[sheet], code), "before\nafter");
});

test("literal CR in a requested formula roundtrips or is refused without output", async (t) => {
	const { dir, path, output } = await setup(t);
	const formula = '=IF(1,"a\rb","c")';
	const parts = await editOrRefuse(dir, path, output, /Invalid XLSX formula request/, formula);
	if (!parts) return;
	const text = inspect(parts[sheet], 'print(json.dumps(r.find(".//"+M+"c[@r=\'B1\']/"+M+"f").text))');
	assert.equal(text, formula.slice(1));
});

test("output through a directory symlink to the source is refused", async (t) => {
	const { dir, path } = await setup(t);
	const original = await readFile(path);
	const alias = join(dir, "alias");
	await symlink(dir, alias, "dir");
	let rejected = false;
	try { await executeFiller(request(path, join(alias, "source.xlsx")), dir); } catch { rejected = true; }
	assert.deepEqual(await readFile(path), original, "source must remain unchanged even through aliases");
	assert.ok(rejected, "alias must be refused");
});

const variants: Array<[string, string, (xml: string) => string]> = [
	["foreign default namespace in workbook", workbook, (xml) => xml.replace('<workbook xmlns=', '<m:workbook xmlns:m=').replace('</workbook>', '</m:workbook>').replace('<sheets>', '<m:sheets>').replace('</sheets>', '</m:sheets>').replace('<sheet ', '<m:sheet ').replace('xmlns:r=', 'xmlns="urn:foreign" xmlns:r=')],
	["foreign default namespace in worksheet", sheet, (xml) => xml.replace('<worksheet xmlns=', '<m:worksheet xmlns:m=').replace('</worksheet>', '</m:worksheet>').replace('<sheetData>', '<m:sheetData>').replace('</sheetData>', '</m:sheetData>').replace('<row ', '<m:row ').replace('</row>', '</m:row>').replace(/<c /g, '<m:c ').replace(/<\/c>/g, '</m:c>').replace('<is><t>', '<m:is><m:t>').replace('</t></is>', '</m:t></m:is>').replace('<v>', '<m:v>').replace('</v>', '</m:v>').replace('><m:sheetData>', ' xmlns="urn:foreign"><m:sheetData>')],
	["mixed foreign content in workbook", workbook, (xml) => xml.replace('</workbook>', '<x:marker xmlns:x="urn:foreign">alpha<x:child/>omega</x:marker></workbook>')],
	["mixed foreign content in worksheet", sheet, (xml) => xml.replace('</worksheet>', '<x:marker xmlns:x="urn:foreign">alpha<x:child/>omega</x:marker></worksheet>')],
	["alternate content in workbook", workbook, (xml) => xml.replace('</workbook>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:foreign"><mc:Choice Requires="x"><x:marker>choice</x:marker></mc:Choice><mc:Fallback><x:marker>fallback</x:marker></mc:Fallback></mc:AlternateContent></workbook>')],
	["alternate content in worksheet", sheet, (xml) => xml.replace('</worksheet>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:foreign"><mc:Choice Requires="x"><x:marker>choice</x:marker></mc:Choice><mc:Fallback><x:marker>fallback</x:marker></mc:Fallback></mc:AlternateContent></worksheet>')],
];
for (const [name, part, mutate] of variants) {
	test(`${name} is preserved with correct namespaces or refused`, async (t) => {
		const input = fixture();
		input[part] = enc(mutate(dec(input[part])));
		inspect(input[part], 'print(json.dumps(r.tag))'); // Ensure the synthetic input is namespace-well-formed.
		const { dir, path, output } = await setup(t, input);
		const parts = await editOrRefuse(dir, path, output, /Unsupported XLSX formula (XML namespace|mixed XML content)/);
		if (!parts) return;
		const code = 'def tree(e):\n return [e.tag, sorted(e.attrib.items()), e.text, [[tree(c), c.tail] for c in e]]\nprint(json.dumps([tree(e) for e in r if e.tag in ("{urn:foreign}marker", "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent")]))';
		assert.deepEqual(inspect(parts[part], code), inspect(input[part], code));
		const calc = inspect(parts[workbook], 'print(json.dumps([e.tag for e in r if e.tag.endswith("}calcPr") or e.tag == "calcPr"]))');
		assert.deepEqual(calc, [`{${main}}calcPr`]);
		const dimensions = inspect(parts[sheet], 'print(json.dumps([e.tag for e in r if e.tag.endswith("}dimension") or e.tag == "dimension"]))');
		assert.ok(Array.isArray(dimensions));
		assert.ok(dimensions.every((tag: string) => tag === `{${main}}dimension`));
	});
}
