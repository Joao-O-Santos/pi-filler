import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";
import { unzipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const fods = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
 office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:automatic-styles><style:style style:name="bold" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles>
 <office:body><office:spreadsheet>
 ${["Alpha", "Beta"]
		.map(
			(name) => `<table:table table:name="${name}"><table:table-row>
 <table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
 <table:table-cell table:style-name="bold" table:formula="of:=[.A1]*2" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell>
 <table:table-cell table:formula="of:=[.B1]+1" office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>
 <table:table-cell table:style-name="bold" office:value-type="float" office:value="9"><text:p>9</text:p></table:table-cell>
 </table:table-row></table:table>`,
		)
		.join("")}
 </office:spreadsheet></office:body></office:document>`;

type Cell = {
	formula: string | null;
	value: string | null;
	hasValue: boolean;
	style: string | null;
};
type Inspection = {
	sheets: Record<string, { part: string; cells: Record<string, Cell> }>;
	calc: Record<string, string>;
};

// Independently resolve sheet relationships and inspect ZIP/XML with Python's standard library.
function inspect(path: string): Inspection {
	const script = `import json, posixpath, sys, zipfile, xml.etree.ElementTree as E
z = zipfile.ZipFile(sys.argv[1])
n = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
p = '{http://schemas.openxmlformats.org/package/2006/relationships}'
w = E.fromstring(z.read('xl/workbook.xml'))
rels = E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
targets = {x.get('Id'): x.get('Target') for x in rels.findall(p+'Relationship')}
sheets = {}
for s in w.find(n+'sheets'):
    target = targets[s.get(r+'id')]
    part = posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join('xl', target))
    root = E.fromstring(z.read(part))
    cells = {}
    for c in root.findall('.//'+n+'c'):
        f, v = c.find(n+'f'), c.find(n+'v')
        cells[c.get('r')] = {'formula': None if f is None else f.text, 'value': None if v is None else v.text, 'hasValue': v is not None, 'style': c.get('s')}
    sheets[s.get('name')] = {'part': part, 'cells': cells}
calc = w.find(n+'calcPr')
print(json.dumps({'sheets': sheets, 'calc': {} if calc is None else calc.attrib}))`;
	return JSON.parse(execFileSync("python3", ["-c", script, path], { encoding: "utf8" }));
}

test("real LibreOffice multi-sheet formulas survive export; dependent caches are observed", async (t) => {
	const probe = spawnSync("libreoffice", ["--version"], { encoding: "utf8", timeout: 15000 });
	if ((probe.error as NodeJS.ErrnoException | undefined)?.code === "ENOENT") {
		t.skip("LibreOffice is not installed on PATH");
		return;
	}
	if (probe.error) throw probe.error;
	assert.equal(probe.status, 0, probe.stderr);
	const dir = await mkdtemp(join(tmpdir(), "xlsx-multisheet-lo-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const source = join(dir, "source.fods");
	const input = join(dir, "source.xlsx");
	const output = join(dir, "edited.xlsx");
	const exported = join(dir, "exported");
	await mkdir(exported);
	await writeFile(source, fods);
	const convert = (file: string, outdir: string, profile: string) =>
		execFileSync(
			"libreoffice",
			[
				`-env:UserInstallation=${pathToFileURL(join(dir, profile)).href}`,
				"--headless",
				"--convert-to",
				"xlsx",
				"--outdir",
				outdir,
				file,
			],
			{ timeout: 30000 },
		);
	convert(source, dir, "create-profile");
	const originalBytes = await readFile(input);
	const before = inspect(input);
	for (const name of ["Alpha", "Beta"]) {
		const c = before.sheets[name].cells;
		assert.equal(c.B1.formula, "A1*2");
		assert.equal(c.B1.value, "6");
		assert.equal(c.C1.value, "7");
		assert.ok(c.B1.style);
		assert.ok(c.D1.style);
	}
	assert.notEqual(before.sheets.Alpha.part, before.sheets.Beta.part);
	const result = await executeFiller(
		{
			format: "xlsx",
			action: "patch",
			path: input,
			output,
			xlsx_formula_sheets: [
				{
					sheet: "alpha",
					formulas: [
						{ cell: "B1", formula: "=A1*4" },
						{ cell: "D1", formula: "=A1+10" },
					],
				},
				{
					sheet: "BETA",
					formulas: [
						{ cell: "B1", formula: "=A1*5" },
						{ cell: "D1", formula: "=A1+20" },
					],
				},
			],
		},
		dir,
	);
	assert.equal(result.details.formulaCount, 4);
	assert.deepEqual(await readFile(input), originalBytes);
	const after = inspect(output);
	const oldParts = unzipSync(originalBytes);
	const newParts = unzipSync(await readFile(output));
	assert.deepEqual(Object.keys(newParts).sort(), Object.keys(oldParts).sort());
	const editedParts = ["xl/workbook.xml", before.sheets.Alpha.part, before.sheets.Beta.part];
	for (const part of Object.keys(oldParts))
		if (!editedParts.includes(part)) assert.deepEqual(newParts[part], oldParts[part], part);
	for (const [name, b, d] of [
		["Alpha", "A1*4", "A1+10"],
		["Beta", "A1*5", "A1+20"],
	]) {
		const original = before.sheets[name].cells;
		const c = after.sheets[name].cells;
		assert.equal(after.sheets[name].part, before.sheets[name].part);
		for (const [address, formula] of [
			["B1", b],
			["D1", d],
		]) {
			assert.equal(c[address].formula, formula);
			assert.equal(c[address].hasValue, false);
			assert.equal(c[address].style, original[address].style);
		}
		assert.deepEqual(c.A1, original.A1);
		assert.deepEqual(c.C1, original.C1);
	}
	assert.equal(after.calc.calcMode, "auto");
	assert.equal(after.calc.fullCalcOnLoad, "1");
	assert.equal(after.calc.forceFullCalc, "1");
	convert(output, exported, "reopen-profile");
	const reopened = inspect(join(exported, "edited.xlsx"));
	for (const name of ["Alpha", "Beta"]) {
		for (const address of ["B1", "D1"]) {
			assert.equal(
				reopened.sheets[name].cells[address].formula,
				after.sheets[name].cells[address].formula,
			);
			assert.ok(reopened.sheets[name].cells[address].style);
		}
		assert.equal(reopened.sheets[name].cells.C1.formula, "B1+1");
	}
	t.diagnostic(
		`${probe.stdout.trim()}; reopened cells: ${JSON.stringify(Object.fromEntries(Object.entries(reopened.sheets).map(([name, sheet]) => [name, sheet.cells])))}`,
	);
});
