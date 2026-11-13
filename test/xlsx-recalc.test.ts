import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
const WORKBOOK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/><sheet name="T 2" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029"/></workbook>`;
const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`;
// Every cached value is deliberately stale (0), so a successful run must replace it.
const SHEET1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>0</v></c><c r="C1"><f>A1&gt;1</f><v>0</v></c><c r="D1"><f>"x"&amp;A1</f><v></v></c><c r="E1"><f>1/0</f><v>0</v></c></row></sheetData></worksheet>`;
const SHEET2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>S!B1+1</f><v>0</v></c><c r="B1"><f>'T 2'!A1*2</f><v>0</v></c><c r="C1"><f>TRUE()</f><v>0</v></c></row></sheetData></worksheet>`;

function workbook(extra: Record<string, Uint8Array> = {}): Uint8Array {
	return zipSync({
		"[Content_Types].xml": new TextEncoder().encode(CT),
		"_rels/.rels": new TextEncoder().encode(ROOT_RELS),
		"xl/workbook.xml": new TextEncoder().encode(WORKBOOK),
		"xl/_rels/workbook.xml.rels": new TextEncoder().encode(WORKBOOK_RELS),
		"xl/worksheets/sheet1.xml": new TextEncoder().encode(SHEET1),
		"xl/worksheets/sheet2.xml": new TextEncoder().encode(SHEET2),
		...extra,
	});
}

function cells(bytes: Uint8Array, part: string): Record<string, string> {
	const xml = new TextDecoder().decode(unzipSync(bytes)[part]);
	const result: Record<string, string> = {};
	// Whole elements, so the t attribute on <c> is visible to assertions.
	for (const match of xml.matchAll(/<c r="([A-Z]+[0-9]+)"[^>]*>.*?<\/c>/g)) {
		result[match[1]] = match[0];
	}
	return result;
}

const libreoffice = spawnSync("libreoffice", ["--version"], { encoding: "utf8", timeout: 15000 });
const skip =
	(libreoffice.error as NodeJS.ErrnoException | undefined)?.code === "ENOENT"
		? "LibreOffice is not installed on PATH"
		: false;

test("recalculation replaces stale caches for cross-sheet, text, boolean and error results", {
	skip,
	timeout: 300_000,
}, async () => {
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-recalc-test-"));
	try {
		const input = join(dir, "stale.xlsx");
		const output = join(dir, "recalc.xlsx");
		const bytes = workbook();
		await writeFile(input, bytes);
		const digest = createHash("sha256").update(bytes).digest("hex");
		const result = await executeFiller(
			{ format: "xlsx", action: "patch", path: input, output, xlsx_recalculate: true },
			dir,
		);
		const details = result.details as Record<string, unknown>;
		assert.equal(details.written, true);
		assert.equal(details.formulaCount, 7);
		assert.deepEqual(details.changedParts, [
			"xl/worksheets/sheet1.xml",
			"xl/worksheets/sheet2.xml",
		]);
		assert.doesNotMatch(result.text, /x2|#DIV/, "values must not be returned");

		const sheet1 = cells(await readFile(output), "xl/worksheets/sheet1.xml");
		assert.match(sheet1.B1, /<f>A1\*3<\/f><v>6<\/v>/);
		assert.doesNotMatch(sheet1.B1, /t="/);
		assert.match(sheet1.C1, /t="b"/);
		assert.match(sheet1.D1, /t="str"/);
		assert.match(sheet1.D1, /<v>x2<\/v>/);
		assert.match(sheet1.E1, /t="e"/);
		assert.match(sheet1.E1, /<v>#DIV\/0!<\/v>/);
		const sheet2 = cells(await readFile(output), "xl/worksheets/sheet2.xml");
		assert.match(sheet2.A1, /<v>7<\/v>/);
		assert.match(sheet2.B1, /<v>14<\/v>/);
		assert.match(sheet2.C1, /t="b"/);
		// Formula text and the source are unchanged.
		assert.match(sheet2.A1, /<f[^>]*>S!B1\+1<\/f>/);
		assert.equal(
			createHash("sha256")
				.update(await readFile(input))
				.digest("hex"),
			digest,
		);
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});

test("dry run validates recalculation without writing", { skip, timeout: 300_000 }, async () => {
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-recalc-dry-"));
	try {
		const input = join(dir, "stale.xlsx");
		const output = join(dir, "dry.xlsx");
		await writeFile(input, workbook());
		const result = await executeFiller(
			{
				format: "xlsx",
				action: "patch",
				path: input,
				output,
				xlsx_recalculate: true,
				dry_run: true,
			},
			dir,
		);
		assert.equal((result.details as Record<string, unknown>).written, false);
		await assert.rejects(stat(output), { code: "ENOENT" });
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});

test("recalculation refuses external links, same-path output and bad arguments", async () => {
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-recalc-refuse-"));
	try {
		const input = join(dir, "external.xlsx");
		const output = join(dir, "out.xlsx");
		const link = new TextEncoder().encode("<externalLink/>");
		await writeFile(input, workbook({ "xl/externalLinks/externalLink1.xml": link }));
		await assert.rejects(
			executeFiller(
				{ format: "xlsx", action: "patch", path: input, output, xlsx_recalculate: true },
				dir,
			),
			/external workbook links/,
		);
		await assert.rejects(stat(output), { code: "ENOENT" });

		const plain = join(dir, "plain.xlsx");
		await writeFile(plain, workbook());
		await writeFile(output, "existing");
		// Dry runs validate but must leave an existing output untouched.
		await executeFiller(
			{
				format: "xlsx",
				action: "patch",
				path: plain,
				output,
				xlsx_recalculate: true,
				dry_run: true,
			},
			dir,
		);
		assert.equal(await readFile(output, "utf8"), "existing");
		await assert.rejects(
			executeFiller(
				{ format: "xlsx", action: "patch", path: plain, output: plain, xlsx_recalculate: true },
				dir,
			),
			/distinct|must differ/,
		);
		await assert.rejects(
			executeFiller(
				{
					format: "xlsx",
					action: "patch",
					path: plain,
					output: join(dir, "x.xlsx"),
					xlsx_recalculate: false as unknown as true,
				},
				dir,
			),
			/xlsx_recalculate must be true/,
		);
		await assert.rejects(
			executeFiller(
				{
					format: "xlsx",
					action: "patch",
					path: plain,
					output: join(dir, "y.xlsx"),
					xlsx_recalculate: true,
					sheet: "S",
				},
				dir,
			),
		);
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});
