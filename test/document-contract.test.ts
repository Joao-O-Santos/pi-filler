import assert from "node:assert/strict";
import { rm } from "node:fs/promises";
import { join } from "node:path";
import test from "node:test";
import extension, { executeFiller } from "../src/index.js";
import { directory, documentAt, fixture, p } from "./docx-edit-fixture.js";

test("new operation contract rejects irrelevant fields before file access and advertises units/limitations", async () => {
	const bad = [
		{ format: "docx", action: "export", output: "x.pdf" },
		{ format: "pdf", action: "export", output: "x.pdf", view: "image" },
		{ format: "pdf", action: "export", output: "x.pdf", dry_run: false },
		{
			format: "pdf",
			action: "compose",
			output: "x.pdf",
			docx_compose: { body_path: "b", replace_from: "x" },
		},
		{
			format: "docx",
			action: "compose",
			output: "x",
			docx_compose: { body_path: "b", replace_from: "x" },
			overwrite_original: true,
		},
		{
			format: "docx",
			action: "patch",
			output: "x",
			patches: { clearCoreMetadata: true },
			docx_layout: { paragraphs: [{ from: 1, size: 11 }] },
		},
		{ format: "docx", action: "read", view: "text", structure_limit: 1 },
		{ format: "pdf", action: "read", view: "structure" },
	];
	for (const value of bad)
		await assert.rejects(
			executeFiller({ path: "/nonexistent", ...value } as never, "/"),
			(error) => !String(error).includes("ENOENT"),
		);
	let guidelines: string[] = [];
	extension({
		registerTool(tool: { promptGuidelines: string[] }) {
			guidelines = tool.promptGuidelines;
		},
	} as never);
	for (const requirement of [
		/action export/,
		/docx_compose/,
		/docx_layout/,
		/240ths/,
		/effective inherited appearance/,
		/vertical\/legacy merges/,
		/overflow/,
	])
		assert.ok(
			guidelines.some((line) => requirement.test(line)),
			String(requirement),
		);
});

test("document operations route structure, layout and composition through the tool", async (t) => {
	const dir = await directory();
	t.after(() => rm(dir, { recursive: true, force: true }));
	const shell = join(dir, "shell.docx"),
		donor = join(dir, "body.docx"),
		shaped = join(dir, "shaped.docx"),
		output = join(dir, "output.docx");
	await fixture(shell, p("Shell") + p("REPLACE"));
	await fixture(donor, p("Body"));
	const audit = await executeFiller(
		{ format: "docx", action: "read", view: "structure", path: shell, structure_limit: 1 },
		dir,
	);
	assert.equal(JSON.parse(audit.text).parts[0].paragraphs.length, 1);
	await executeFiller(
		{
			format: "docx",
			action: "patch",
			path: donor,
			output: shaped,
			docx_layout: { paragraphs: [{ from: 1, size: 12 }] },
		},
		dir,
	);
	assert.match(await documentAt(shaped), /w:sz w:val="24"/);
	const composed = await executeFiller(
		{
			format: "docx",
			action: "compose",
			path: shell,
			output,
			docx_compose: { body_path: shaped, replace_from: "REPLACE" },
		},
		dir,
	);
	assert.equal(composed.details.insertedElements, 1);
	assert.match(await documentAt(output), />Body</);
});
