// biome-ignore-all format: compact synthetic OPC fixtures
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { unzipSync, zipSync } from "fflate";
import { executeFiller } from "../src/index.js";

const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const relns = "http://schemas.openxmlformats.org/package/2006/relationships";
const office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const typens = "http://schemas.openxmlformats.org/package/2006/content-types";
const rels = "xl/_rels/workbook.xml.rels";
const types = "[Content_Types].xml";
const secret = "PRIVATE_CHAIN_SAFETY_79b210";
const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);
type Parts = Record<string, Uint8Array>;

function fixture(): Parts {
	return Object.fromEntries(Object.entries({
		[types]: `<Types xmlns="${typens}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
		"_rels/.rels": `<Relationships xmlns="${relns}"><Relationship Id="office" Type="${office}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml": `<workbook xmlns="${main}" xmlns:r="${office}"><sheets><sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>`,
		[rels]: `<Relationships xmlns="${relns}"><Relationship Id="sheet" Type="${office}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="chain" Type="${office}/calcChain" Target="calcChain.xml"/></Relationships>`,
		"xl/worksheets/sheet1.xml": `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
		"xl/calcChain.xml": `<calcChain xmlns="${main}"><c r="A1" i="1"/></calcChain>`,
	}).map(([part, xml]) => [part, enc(xml)]));
}

async function setup(t: TestContext, parts: Parts) {
	const dir = await mkdtemp(join(tmpdir(), "xlsx-chain-safety-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const path = join(dir, "input.xlsx");
	const output = join(dir, "output.xlsx");
	await writeFile(path, zipSync(parts));
	await writeFile(output, "previous output");
	return { dir, path, output };
}

async function refuse(t: TestContext, parts: Parts, dry_run: boolean) {
	const { dir, path, output } = await setup(t, parts);
	const before = await readFile(path);
	let error: unknown;
	try {
		await executeFiller({ format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell: "A1", formula: `="${secret}"` }, dry_run } as never, dir);
	} catch (caught) { error = caught; }
	assert.deepEqual(await readFile(path), before, "source must remain byte-identical");
	assert.ok(error, "unsafe calculation-chain packaging must be refused");
	assert.ok(!String(error).includes(secret), "error must not disclose formula contents");
	assert.equal(await readFile(output, "utf8"), "previous output", "refusal must not publish");
}

for (const dry_run of [false, true]) {
	test(`mixed-prefix incoming relationship must not survive chain deletion (dry_run=${dry_run})`, async (t) => {
		const parts = fixture();
		// Both elements are genuine OPC relationships in the SAME namespace.
		// directObjects currently visits only the first spelling of Relationship.
		parts["_rels/.rels"] = enc(dec(parts["_rels/.rels"]).replace("</Relationships>", `<p:Relationship xmlns:p="${relns}" Id="also-chain" Type="urn:custom" Target="xl/calcChain.xml"/></Relationships>`));
		await refuse(t, parts, dry_run);
	});

	for (const part of [rels, types]) {
		test(`foreign-namespace chain metadata must be refused in ${part} (dry_run=${dry_run})`, async (t) => {
			const parts = fixture();
			const marker = part === rels ? '<Relationship Id="chain"' : '<Override PartName="/xl/calcChain.xml"';
			parts[part] = enc(dec(parts[part]).replace(marker, `${marker} xmlns="urn:not-opc"`));
			await refuse(t, parts, dry_run);
		});
	}

	for (const withChain of [false, true]) {
		test(`external URL is not an incoming package edge (chain=${withChain}, dry_run=${dry_run})`, async (t) => {
			const parts = fixture();
			if (!withChain) {
				delete parts["xl/calcChain.xml"];
				parts[rels] = enc(dec(parts[rels]).replace(/<Relationship Id="chain"[^>]*\/>/, ""));
				parts[types] = enc(dec(parts[types]).replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, ""));
			}
			// This points OUTSIDE the package, not to its similarly named part.
			parts["_rels/.rels"] = enc(dec(parts["_rels/.rels"]).replace("</Relationships>", `<Relationship Id="external" Type="${office}/hyperlink" TargetMode="External" Target="xl/calcChain.xml"/></Relationships>`));
			const { dir, path, output } = await setup(t, parts);
			const before = await readFile(path);
			const result = await executeFiller({ format: "xlsx", action: "patch", path, output, sheet: "Data", xlsx_formula: { cell: "A1", formula: `="${secret}"` }, dry_run } as never, dir);
			assert.ok(!JSON.stringify(result).includes(secret));
			assert.deepEqual(await readFile(path), before);
			if (dry_run) assert.equal(await readFile(output, "utf8"), "previous output");
			else assert.deepEqual(unzipSync(await readFile(output))["_rels/.rels"], parts["_rels/.rels"]);
		});
	}

	for (const badByte of [0, 255]) {
		test(`metadata byte 0x${badByte.toString(16)} must not be silently rewritten (dry_run=${dry_run})`, async (t) => {
			const parts = fixture();
			// Unrelated external relationship is preserved while metadata is rebuilt.
			parts[rels] = enc(dec(parts[rels]).replace("</Relationships>", `<Relationship Id="external" Type="${office}/hyperlink" TargetMode="External" Target="https://example.invalid/${secret}"/></Relationships>`));
			const index = Buffer.from(parts[rels]).indexOf(secret);
			assert.notEqual(index, -1);
			parts[rels][index] = badByte;
			await refuse(t, parts, dry_run);
		});
	}
}
