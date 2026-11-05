import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { executeFiller, type FillerInput, fillerSchema } from "../src/index.js";
import { readDocxText, readPptxText, searchDocx, searchPptx, writeDocx } from "../src/pandoc.js";

test("PPTX extraction remains GFM without revision flags", async (t) => {
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-pptx-flags-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const executable = join(dir, "pandoc");
	const log = join(dir, "args");
	await writeFile(executable, '#!/bin/sh\nprintf "%s\\n" "$@" > "$LOG"\nprintf "slide text\\n"\n');
	await chmod(executable, 0o755);
	const env = { ...process.env, PATH: `${dir}:${process.env.PATH ?? ""}`, LOG: log };
	for (const run of [
		() => readPptxText(join(dir, "a.pptx"), { env }),
		() => searchPptx(join(dir, "a.pptx"), "slide", { env }),
	]) {
		await run();
		const args = await readFile(log, "utf8");
		assert.match(args, /^--from=pptx\n--to=gfm\n/);
		assert.doesNotMatch(args, /track-changes|markdown-smart/);
	}
});

const probe = spawnSync("pandoc", ["--version"]);
const absent = probe.error && "code" in probe.error && probe.error.code === "ENOENT";

test("real Pandoc preserves DOCX revisions and anchored comments", {
	skip: absent ? "pandoc absent" : false,
}, async (t) => {
	assert.equal(probe.status, 0);
	const dir = await mkdtemp(join(tmpdir(), "pi-filler-review-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const source = join(dir, "base.md");
	const base = join(dir, "base.docx");
	const input = join(dir, "review.docx");
	const markdown = join(dir, "review.md");
	const output = join(dir, "roundtrip.docx");
	await writeFile(source, "Placeholder\n");
	await writeDocx(source, base);
	const parts = unzipSync(await readFile(base));
	const document = strFromU8(parts["word/document.xml"]);
	const review =
		'<w:commentRangeStart w:id="5"/><w:r><w:t>Anchor</w:t></w:r><w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:ins w:id="7" w:author="Alice" w:date="2024-01-02T03:04:05Z"><w:r><w:t>Inserted</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> </w:t></w:r><w:del w:id="8" w:author="Bob" w:date="2024-01-03T03:04:05Z"><w:r><w:delText>Deleted</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> &quot;quoted&quot; 1--2 001.230</w:t></w:r>';
	assert.match(document, /<w:r><w:t[^>]*>Placeholder<\/w:t><\/w:r>/);
	parts["word/document.xml"] = strToU8(
		document.replace(/<w:r><w:t[^>]*>Placeholder<\/w:t><\/w:r>/, review),
	);
	parts["word/comments.xml"] = strToU8(
		'<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="5" w:author="Carol" w:date="2024-01-04T03:04:05Z"><w:p><w:r><w:t>Review note</w:t></w:r></w:p></w:comment></w:comments>',
	);
	const rels = strFromU8(parts["word/_rels/document.xml.rels"]);
	if (!rels.includes('relationships/comments"'))
		parts["word/_rels/document.xml.rels"] = strToU8(
			rels.replace(
				"</Relationships>",
				'<Relationship Id="reviewComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
			),
		);
	const types = strFromU8(parts["[Content_Types].xml"]);
	if (!types.includes("/word/comments.xml"))
		parts["[Content_Types].xml"] = strToU8(
			types.replace(
				"</Types>",
				'<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
			),
		);
	await writeFile(input, zipSync(parts));
	const all = (await readDocxText(input, { trackChanges: "all" })).text;
	const accepted = (await readDocxText(input)).text;
	const rejected = (await readDocxText(input, { trackChanges: "reject" })).text;
	assert.match(all, /\[Inserted\]\{\.insertion[^}]*author="Alice"/);
	assert.match(all, /\[Deleted\]\{\.deletion[^}]*author="Bob"/);
	assert.match(all, /comment-start/);
	assert.match(all, /Review note/);
	assert.match(all, /Carol/);
	assert.match(accepted, /Inserted/);
	assert.doesNotMatch(accepted, /Deleted/);
	assert.match(rejected, /Deleted/);
	assert.doesNotMatch(rejected, /Inserted/);
	assert.equal((await searchDocx(input, "Deleted", { trackChanges: "reject" })).matchCount, 1);
	assert.match(all, /"quoted" 1--2/);
	for (const [mode, expected] of [
		["accept", accepted],
		["reject", rejected],
		["all", all],
	] as const) {
		const result = await executeFiller(
			{ format: "docx", action: "read", view: "text", path: input, track_changes: mode },
			dir,
		);
		assert.equal(result.text, expected);
	}
	const searched = await executeFiller(
		{ format: "docx", action: "search", path: input, query: "Deleted", track_changes: "reject" },
		dir,
	);
	assert.equal(searched.details.matchCount, 1);
	await writeFile(markdown, all);
	await writeDocx(markdown, output);
	const roundtrip = unzipSync(await readFile(output));
	const xml = strFromU8(roundtrip["word/document.xml"]);
	assert.match(xml, /<w:ins\b/);
	assert.match(xml, /<w:del\b/);
	assert.match(xml, /<w:delText\b/);
	assert.match(xml, /<w:commentRangeStart\b/);
	assert.match(xml, /<w:commentRangeEnd\b/);
	assert.ok(roundtrip["word/comments.xml"]);
	assert.match(strFromU8(roundtrip["word/comments.xml"]), /<w:comment\b[^>]*[\s\S]*Review note/);
	const reread = (await readDocxText(output, { trackChanges: "all" })).text;
	assert.match(reread, /"quoted" 1--2 001\.230/);
	assert.match(reread, /\[Review note\]\{\.comment-start[^}]*\}Anchor\[\]\{\.comment-end[^}]*\}/);
	assert.match(reread, /\[Inserted\]\{\.insertion[^}]*author="Alice"[^}]*2024-01-02T03:04:05Z/);
	assert.match(reread, /\[Deleted\]\{\.deletion[^}]*author="Bob"[^}]*2024-01-03T03:04:05Z/);
	assert.match(reread, /Review note.*comment-start[^}]*Carol/);
	assert.equal((await readDocxText(output)).text, accepted);
	assert.equal((await readDocxText(output, { trackChanges: "reject" })).text, rejected);
});

test("revision-view schema is explicit and restricted to DOCX text reads/searches", async () => {
	const schema = fillerSchema.properties.track_changes as unknown as Record<string, unknown>;
	assert.equal(schema.default, "accept");
	assert.deepEqual(schema.enum, ["accept", "reject", "all"]);
	for (const input of [
		{ format: "docx", action: "read", view: "formatting" },
		{ format: "docx", action: "read", view: "image", output: "render" },
		{ format: "docx", action: "write", output: "out.docx" },
		{ format: "pptx", action: "read", view: "text" },
		{ format: "pdf", action: "read", view: "text" },
		{ format: "xlsx", action: "read", view: "structure" },
	] as const) {
		await assert.rejects(
			() =>
				executeFiller(
					{ ...input, path: "missing.file", track_changes: "all" } as FillerInput,
					"/tmp",
				),
			/does not accept track_changes/,
		);
	}
});
