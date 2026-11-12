import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { constants } from "node:fs";
import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { delimiter, join } from "node:path";
import test from "node:test";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { executeFiller, type FillerInput } from "../src/index.js";

const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const review = `<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>commented</w:t></w:r><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r><w:ins w:id="7" w:author="Prior" w:date="2024-01-02T00:00:00Z"><w:r><w:t>earlier insertion</w:t></w:r></w:ins><w:del w:id="8" w:author="Prior" w:date="2024-01-03T00:00:00Z"><w:r><w:delText>earlier deletion</w:delText></w:r></w:del></w:p>`;
const comments = `<w:comments xmlns:w="${w}"><w:comment w:id="4" w:author="Reviewer"><w:p><w:r><w:t>Synthetic note</w:t></w:r></w:p></w:comment></w:comments>`;

function packageBytes() {
	const parts: Record<string, string> = {
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
		"_rels/.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`,
		"word/document.xml": `<w:document xmlns:w="${w}"><w:body>${review}<w:p><w:r><w:t>Replace this sentence.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"word/comments.xml": comments,
	};
	return zipSync(
		Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
	);
}

async function available() {
	for (const dir of (process.env.PATH ?? "").split(delimiter)) {
		try {
			await access(join(dir || ".", "docxquisite"), constants.X_OK);
			return true;
		} catch {
			// Try the next PATH entry.
		}
	}
	return false;
}

test("real toolkit tracked replacement preserves independent review markup and refuses unsafe edits", async (t) => {
	if (!(await available())) {
		t.skip("docxquisite executable is not available on PATH");
		return;
	}
	const dir = await mkdtemp(join(tmpdir(), "filler-real-toolkit-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const source = join(dir, "source.docx");
	const bytes = packageBytes();
	await writeFile(source, bytes);
	const hash = createHash("sha256").update(bytes).digest("hex");
	const call = (output: string, old: string, expected_sha256: string) =>
		executeFiller(
			{
				format: "docx",
				action: "patch",
				path: "@source.docx",
				output: `@${output}`,
				tracked_replace: { old, new: "Replacement text.", author: "Test Editor", expected_sha256 },
			} as FillerInput,
			dir,
		);
	await call("output.docx", "Replace this sentence.", hash);
	assert.deepEqual(await readFile(source), Buffer.from(bytes));
	const originalParts = unzipSync(bytes);
	const parts = unzipSync(await readFile(join(dir, "output.docx")));
	assert.deepEqual(Object.keys(parts).sort(), Object.keys(originalParts).sort());
	for (const name of Object.keys(parts).filter((name) => name !== "word/document.xml")) {
		assert.deepEqual(parts[name], originalParts[name], name);
	}
	const xml = strFromU8(parts["word/document.xml"]);
	assert.equal(XMLValidator.validate(xml), true);
	const parsed = new XMLParser({ ignoreAttributes: false, removeNSPrefix: true }).parse(xml);
	const paragraphs = parsed.document.body.p;
	assert.equal(paragraphs.length, 2);
	assert.equal(paragraphs[0].commentRangeStart["@_id"], "4");
	assert.equal(paragraphs[0].commentRangeEnd["@_id"], "4");
	assert.equal(paragraphs[0].ins.r.t, "earlier insertion");
	assert.equal(paragraphs[0].del.r.delText, "earlier deletion");
	assert.equal(paragraphs[1].del.r.delText["#text"], "Replace this sentence.");
	assert.equal(paragraphs[1].ins.r.t["#text"], "Replacement text.");
	assert.equal(paragraphs[1].ins["@_author"], "Test Editor");
	assert.equal(paragraphs[1].del["@_author"], "Test Editor");
	const ids = [
		paragraphs[0].ins["@_id"],
		paragraphs[0].del["@_id"],
		paragraphs[1].ins["@_id"],
		paragraphs[1].del["@_id"],
	];
	assert.equal(new Set(ids).size, 4);

	await assert.rejects(call("stale.docx", "Replace this sentence.", "0".repeat(64)));
	await assert.rejects(readFile(join(dir, "stale.docx")), { code: "ENOENT" });
	await assert.rejects(call("overlap.docx", "commented", hash));
	await assert.rejects(readFile(join(dir, "overlap.docx")), { code: "ENOENT" });
	assert.deepEqual(await readFile(source), Buffer.from(bytes));
});
