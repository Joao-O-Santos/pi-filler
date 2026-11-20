import assert from "node:assert/strict";
import { chmod, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { PDFDocument } from "@cantoo/pdf-lib";
import { exportOfficePdf } from "../src/office-export.js";

async function setup(t: test.TestContext) {
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-export-test-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const input = join(dir, "input.docx"),
		output = join(dir, "result.pdf"),
		fixture = join(dir, "fixture.pdf");
	await writeFile(input, "unchanged source");
	await writeFile(output, "prior output");
	const doc = await PDFDocument.create();
	doc.addPage();
	doc.addPage();
	await writeFile(fixture, await doc.save());
	const script = join(dir, "libreoffice");
	await writeFile(
		script,
		`#!/bin/sh\nif [ "$1" = "--version" ]; then\n if [ "$FAIL_VERSION" = 1 ]; then exit 2; fi\n if [ "$EMPTY_VERSION" != 1 ]; then echo "LibreOffice test-backend"; fi\n exit 0\nfi\nprintf '%s\\n' "$@" > "$EXPORT_LOG"\nif [ "$FAIL_EXPORT" = 1 ]; then echo fail >&2; exit 2; fi\nif [ "$SLOW_EXPORT" = 1 ]; then exec sleep 5; fi\nif [ "$NO_EXPORT" = 1 ]; then exit 0; fi\nif [ "$BAD_EXPORT" = 1 ]; then echo invalid > source.pdf; exit 0; fi\ncp "$PDF_FIXTURE" source.pdf\n`,
	);
	await chmod(script, 0o755);
	const env = {
		...process.env,
		PATH: `${dir}:${process.env.PATH}`,
		PDF_FIXTURE: fixture,
		EXPORT_LOG: join(dir, "args"),
	};
	return { dir, input, output, env };
}

test("PDF export validates PDF, reports page count/version, isolates profile and atomically replaces output", async (t) => {
	const { dir, input, output, env } = await setup(t);
	const result = await exportOfficePdf(input, output, { env });
	assert.equal(result.pageCount, 2);
	assert.equal(result.backend.version, "LibreOffice test-backend");
	assert.match(result.warnings[0], /not Microsoft Office/);
	assert.equal((await PDFDocument.load(await readFile(output))).getPageCount(), 2);
	assert.equal(await readFile(input, "utf8"), "unchanged source");
	const args = await readFile(env.EXPORT_LOG, "utf8");
	assert.match(args, /-env:UserInstallation=file:/);
	const temporary = args.match(/--outdir\n([^\n]+)/)?.[1];
	assert.ok(temporary);
	await assert.rejects(readdir(temporary), /ENOENT/);
	assert.ok(!(await readdir(dir)).some((name) => name.startsWith(".pi-filler")));
});

test("PDF export preserves old output for missing/invalid products and backend/version failures", async (t) => {
	const { input, output, env } = await setup(t);
	for (const variable of [
		"FAIL_VERSION",
		"EMPTY_VERSION",
		"FAIL_EXPORT",
		"NO_EXPORT",
		"BAD_EXPORT",
	]) {
		await assert.rejects(exportOfficePdf(input, output, { env: { ...env, [variable]: "1" } }));
		assert.equal(await readFile(output, "utf8"), "prior output");
	}
	for (const [source, dest] of [
		[input, input],
		[input, output.replace(".pdf", ".docx")],
		[input.replace(".docx", ".txt"), output],
	])
		await assert.rejects(exportOfficePdf(source, dest, { env }));
});

test("PDF export cancellation and timeout do not publish", async (t) => {
	const { input, output, env } = await setup(t);
	const controller = new AbortController();
	controller.abort();
	await assert.rejects(
		exportOfficePdf(input, output, { env, signal: controller.signal }),
		/aborted/i,
	);
	await assert.rejects(
		exportOfficePdf(input, output, { env: { ...env, SLOW_EXPORT: "1" }, timeoutMs: 50 }),
		/timed out/,
	);
	assert.equal(await readFile(output, "utf8"), "prior output");
});
