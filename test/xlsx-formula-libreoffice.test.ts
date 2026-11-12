import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
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
 xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
 office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:body><office:spreadsheet><table:table table:name="Data"><table:table-row>
 <table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
 <table:table-cell table:formula="of:=[.A1]*2" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell>
 <table:table-cell table:formula="of:=[.B1]+1" office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>
 </table:table-row></table:table></office:spreadsheet></office:body>
</office:document>`;

// ElementTree is independent of the XML parser used by the implementation.
function cells(
	xml: Uint8Array,
): Record<string, { formula: string | null; value: string | null; hasValue: boolean }> {
	const script = `import json, sys, xml.etree.ElementTree as E
r = E.fromstring(sys.stdin.buffer.read())
n = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
result = {}
for c in r.findall('.//' + n + 'c'):
    f, v = c.find(n + 'f'), c.find(n + 'v')
    result[c.get('r')] = {'formula': None if f is None else f.text, 'value': None if v is None else v.text, 'hasValue': v is not None}
print(json.dumps(result))`;
	return JSON.parse(execFileSync("python3", ["-c", script], { input: xml, encoding: "utf8" }));
}

test("LibreOffice-produced XLSX formula edit removes only the target cache", async (t) => {
	let office: string | undefined;
	for (const candidate of ["libreoffice", "soffice"]) {
		const probe = spawnSync(candidate, ["--version"], { encoding: "utf8", timeout: 15000 });
		if (probe.error && (probe.error as NodeJS.ErrnoException).code === "ENOENT") continue;
		if (probe.error) throw probe.error;
		assert.equal(probe.status, 0, probe.stderr);
		office = candidate;
		break;
	}
	if (!office) {
		t.skip("LibreOffice is not installed on PATH");
		return;
	}
	const dir = await mkdtemp(join(tmpdir(), "xlsx-formula-lo-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const source = join(dir, "source.fods");
	const workbook = join(dir, "source.xlsx");
	const output = join(dir, "edited.xlsx");
	await writeFile(source, fods);
	execFileSync(
		office,
		[
			`-env:UserInstallation=${pathToFileURL(join(dir, "profile")).href}`,
			"--headless",
			"--convert-to",
			"xlsx",
			"--outdir",
			dir,
			source,
		],
		{ timeout: 30000 },
	);
	const sheet = "xl/worksheets/sheet1.xml";
	const original = cells(unzipSync(await readFile(workbook))[sheet]);
	assert.deepEqual(original.B1, { formula: "A1*2", value: "6", hasValue: true });
	assert.deepEqual(original.C1, { formula: "B1+1", value: "7", hasValue: true });
	await executeFiller(
		{
			format: "xlsx",
			action: "patch",
			path: workbook,
			output,
			sheet: "Data",
			xlsx_formula: { cell: "B1", formula: "=A1*4" },
		} as never,
		dir,
	);
	const edited = cells(unzipSync(await readFile(output))[sheet]);
	assert.deepEqual(edited.B1, { formula: "A1*4", value: null, hasValue: false });
	// This cache remains stale; the patch does not calculate dependent formulas.
	assert.deepEqual(edited.C1, original.C1);
	const batchOutput = join(dir, "batch.xlsx");
	const before = unzipSync(await readFile(workbook));
	await executeFiller(
		{
			format: "xlsx",
			action: "patch",
			path: workbook,
			output: batchOutput,
			sheet: "Data",
			xlsx_formulas: [
				{ cell: "B1", formula: "=A1*5" },
				{ cell: "D4", formula: "=B1+2" },
			],
		},
		dir,
	);
	const after = unzipSync(await readFile(batchOutput));
	const batchCells = cells(after[sheet]);
	assert.deepEqual(batchCells.B1, { formula: "A1*5", value: null, hasValue: false });
	assert.deepEqual(batchCells.D4, { formula: "B1+2", value: null, hasValue: false });
	assert.deepEqual(batchCells.C1, original.C1);
	assert.deepEqual(batchCells.A1, original.A1);
	for (const name of Object.keys(before)) {
		if (![sheet, "xl/workbook.xml"].includes(name)) assert.deepEqual(after[name], before[name]);
	}
});
