import assert from "node:assert/strict";
import { lstat, mkdtemp, readFile, rm, stat, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const enc = new TextEncoder();
const dec = new TextDecoder();

// One worksheet "Data" with A1 = 1 and no formulas, so a formula patch is the only change.
function workbook(): Uint8Array {
	return zipSync({
		"[Content_Types].xml": enc.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
		),
		"_rels/.rels": enc.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
		),
		"xl/workbook.xml": enc.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
		),
		"xl/_rels/workbook.xml.rels": enc.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
		),
		"xl/worksheets/sheet1.xml": enc.encode(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
		),
	});
}

function formulaOf(bytes: Uint8Array): string | undefined {
	const xml = dec.decode(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);
	return /<c r="B1"[^>]*><f>([^<]*)<\/f>/.exec(xml)?.[1];
}

function formulaPatch(path: string, output: string, extra: Record<string, unknown> = {}) {
	return {
		format: "xlsx" as const,
		action: "patch" as const,
		path,
		output,
		sheet: "Data",
		xlsx_formula: { cell: "B1", formula: "=A1+1" },
		...extra,
	};
}

test("XLSX in-place formula patch is refused without overwrite_original and succeeds with it", async () => {
	const dir = await mkdtemp(join(tmpdir(), "filler-inplace-"));
	try {
		const source = join(dir, "book.xlsx");
		const original = workbook();
		await writeFile(source, original);
		await assert.rejects(executeFiller(formulaPatch(source, source), dir), /overwrite_original/);
		assert.deepEqual(new Uint8Array(await readFile(source)), original);
		await executeFiller(formulaPatch(source, source, { overwrite_original: true }), dir);
		assert.equal(formulaOf(new Uint8Array(await readFile(source))), "A1+1");
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});

test("XLSX in-place patch through a symlink alias updates the real file and keeps the alias", async () => {
	const dir = await mkdtemp(join(tmpdir(), "filler-alias-"));
	try {
		const real = join(dir, "real.xlsx");
		const alias = join(dir, "alias.xlsx");
		await writeFile(real, workbook());
		await symlink(real, alias);
		await assert.rejects(executeFiller(formulaPatch(real, alias), dir), /overwrite_original/);
		await executeFiller(formulaPatch(real, alias, { overwrite_original: true }), dir);
		assert.equal(formulaOf(new Uint8Array(await readFile(real))), "A1+1");
		assert.ok((await lstat(alias)).isSymbolicLink());
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});

test("XLSX existing distinct output is replaced on success", async () => {
	const dir = await mkdtemp(join(tmpdir(), "filler-replace-"));
	try {
		const source = join(dir, "book.xlsx");
		const output = join(dir, "out.xlsx");
		await writeFile(source, workbook());
		await writeFile(output, "previous");
		await executeFiller(formulaPatch(source, output), dir);
		assert.equal(formulaOf(new Uint8Array(await readFile(output))), "A1+1");
		assert.ok((await stat(source)).size > 0);
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});

test("overwrite_original is rejected for Markdown-to-DOCX and PPTX writes", async () => {
	const dir = await mkdtemp(join(tmpdir(), "filler-md-"));
	try {
		await writeFile(join(dir, "note.md"), "# Title\n");
		await assert.rejects(
			executeFiller(
				{
					format: "docx",
					action: "write",
					path: join(dir, "note.md"),
					output: join(dir, "note.docx"),
					overwrite_original: true,
				} as never,
				dir,
			),
			/does not accept overwrite_original/,
		);
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});
