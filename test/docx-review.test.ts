import assert from "node:assert/strict";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { strToU8, zipSync } from "fflate";
import { inspectDocxReview } from "../src/docx.js";
import { executeFiller, type FillerInput, fillerSchema } from "../src/index.js";

const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const metadataNs = 'xmlns:a="http://schemas.microsoft.com/office/word/2010/wordml"';
const body = `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Alpha</w:t></w:r>
<w:ins><w:r><w:t> new</w:t></w:r></w:ins><w:del><w:r><w:delText> old</w:delText></w:r></w:del>
<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:r><w:t>Context </w:t></w:r><w:r><w:commentReference w:id="2"/></w:r><w:r><w:t>word</w:t></w:r></w:p>`;
const comments = `<w:comment w:id="0" w:author="A &amp; B" w:date="2026-01-02T03:04:05Z"><w:p a:paraId="AA"><w:r><w:t><![CDATA[First note <001.230>]]></w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Reply author"><w:p a:paraId="BB"><w:r><w:t>Reply</w:t></w:r></w:p><w:p a:paraId="CC"><w:r><w:t>Second paragraph</w:t></w:r></w:p></w:comment>
<w:comment w:id="2"><w:p a:paraId="DD"><w:r><w:t>Point note</w:t></w:r></w:p></w:comment>`;
const state = `<s:commentsEx xmlns:s="urn:extended"><s:commentEx s:paraId="AA" s:done="1"/><s:commentEx s:paraId="CC" s:paraIdParent="AA" s:done="0"/></s:commentsEx>`;

async function fixture(
	t: TestContext,
	options: {
		comments?: string;
		body?: string;
		state?: string;
		modern?: boolean;
		absent?: boolean;
	} = {},
) {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-review-test-"));
	t.after(() => rm(directory, { recursive: true, force: true }));
	const path = join(directory, "input.docx");
	const files: Record<string, Uint8Array> = {
		"[Content_Types].xml": strToU8("<Types/>"),
		"word/document.xml": strToU8(
			`<w:document xmlns:w="${ns}"><w:body>${options.body ?? body}<w:sectPr/></w:body></w:document>`,
		),
	};
	if (!options.absent) {
		files["word/review/notes.xml"] = strToU8(
			`<w:comments xmlns:w="${ns}" ${metadataNs}>${options.comments ?? comments}</w:comments>`,
		);
		files["word/review/state.xml"] = strToU8(options.state ?? state);
		files["word/_rels/document.xml.rels"] = strToU8(
			'<Relationships><Relationship Id="comments" Type="urn:relationships/comments" Target="review/notes.xml"/><Relationship Id="state" Type="urn:relationships/commentsExtended" Target="review/state.xml"/></Relationships>',
		);
	}
	if (options.modern) files["word/people.xml"] = strToU8("<people/>");
	await writeFile(path, zipSync(files));
	return { path, directory };
}

test("review inspection resolves custom parts, ranges, point context and legacy thread state", async (t) => {
	const { path, directory } = await fixture(t);
	const before = await readFile(path);
	const review = await inspectDocxReview(path);
	assert.equal(review.totalComments, 3);
	assert.equal(review.returned, 3);
	assert.equal(review.nextOffset, null);
	assert.deepEqual(review.warnings, []);
	assert.deepEqual(review.comments[0], {
		id: "0",
		author: "A & B",
		date: "2026-01-02T03:04:05Z",
		text: "First note <001.230>",
		textTruncated: false,
		anchor: { kind: "range", text: "Alpha new old", truncated: false },
		resolved: true,
		replyTo: undefined,
	});
	assert.equal(review.comments[1].replyTo, "0");
	assert.equal(review.comments[1].resolved, false);
	assert.equal(review.comments[1].text, "Reply\nSecond paragraph");
	assert.equal(review.comments[1].anchor.kind, "not-found");
	assert.equal(review.comments[2].resolved, undefined);
	assert.deepEqual(review.comments[2].anchor, {
		kind: "point",
		text: "Context word",
		truncated: false,
	});
	const tool = await executeFiller(
		{ format: "docx", action: "read", view: "review", path },
		directory,
	);
	assert.deepEqual(JSON.parse(tool.text), JSON.parse(JSON.stringify(review)));
	assert.deepEqual(tool.details, { totalComments: 3, returned: 3, nextOffset: null });
	assert.deepEqual(await readFile(path), before);
	assert.deepEqual(await readdir(directory), ["input.docx"]);
});

test("review pagination is exact, defaults are explicit and ranges are validated", async (t) => {
	const { path } = await fixture(t);
	const page = await inspectDocxReview(path, 1, 1);
	assert.deepEqual(
		page.comments.map((c) => c.id),
		["1"],
	);
	assert.equal(page.nextOffset, 2);
	assert.equal((await inspectDocxReview(path, 3)).nextOffset, null);
	assert.equal((await inspectDocxReview(path, 999)).returned, 0);
	for (const [offset, limit] of [
		[-1, 1],
		[0.5, 1],
		[0, 0],
		[0, 51],
		[0, 1.5],
	]) {
		await assert.rejects(inspectDocxReview(path, offset, limit), /review_(offset|limit)/);
	}
	const schema = fillerSchema.properties as unknown as Record<string, Record<string, unknown>>;
	assert.equal(schema.review_offset.default, 0);
	assert.equal(schema.review_limit.default, 20);
	assert.equal(schema.review_limit.maximum, 50);
	for (const input of [
		{ format: "docx", action: "read", view: "text" },
		{ format: "docx", action: "search", query: "note" },
		{ format: "pdf", action: "read", view: "text" },
	] as const) {
		await assert.rejects(
			executeFiller({ ...input, path, review_limit: 1 } as FillerInput, "/tmp"),
			/does not accept review_limit/,
		);
	}
	await assert.rejects(
		executeFiller(
			{ format: "docx", action: "read", view: "review", path, track_changes: "all" },
			"/tmp",
		),
		/does not accept track_changes/,
	);
});

test("review output bounds Unicode text and anchors without losing paginated records", async (t) => {
	const big = "文😀".repeat(1600);
	const notes = Array.from(
		{ length: 30 },
		(_, id) => `<w:comment w:id="${id}"><w:p><w:r><w:t>${big}</w:t></w:r></w:p></w:comment>`,
	).join("");
	const { path } = await fixture(t, {
		comments: notes,
		body: `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>${big}</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>`,
		state: "<commentsEx/>",
	});
	const seen: string[] = [];
	let offset = 0;
	do {
		const page = await inspectDocxReview(path, offset, 50);
		assert.ok(Buffer.byteLength(JSON.stringify(page, null, 2)) < 50000);
		assert.ok(page.returned > 0);
		for (const note of page.comments) {
			assert.equal(note.textTruncated, true);
			assert.equal([...note.text].length, 1000);
			assert.doesNotMatch(note.text, /\uFFFD/);
			seen.push(note.id);
		}
		if (offset === 0) {
			assert.equal(page.comments[0].anchor.truncated, true);
			assert.equal([...page.comments[0].anchor.text].length, 200);
			assert.ok(page.returned < 30);
		}
		if (page.nextOffset === null) break;
		assert.ok(page.nextOffset > offset);
		offset = page.nextOffset;
	} while (offset < 30);
	assert.deepEqual(
		seen,
		Array.from({ length: 30 }, (_, id) => String(id)),
	);
});

test("malformed ranges and unsupported metadata are reported rather than inferred", async (t) => {
	const { path } = await fixture(t, {
		body: '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>unfinished</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>',
		state: '<commentsEx><commentEx paraId="AA" done="maybe" paraIdParent="missing"/></commentsEx>',
		modern: true,
	});
	const result = await inspectDocxReview(path);
	assert.equal(result.comments[0].anchor.kind, "unmatched");
	assert.equal(result.comments[0].resolved, undefined);
	assert.equal(result.comments[0].replyTo, undefined);
	assert.match(result.warnings.join(" "), /unmatched comment range end/);
	assert.match(result.warnings.join(" "), /no closing marker/);
	assert.match(result.warnings.join(" "), /resolution value/);
	assert.match(result.warnings.join(" "), /reply parent/);
	assert.match(result.warnings.join(" "), /modern comment metadata/);
});

test("documents without comments return an empty review", async (t) => {
	const { path } = await fixture(t, { absent: true });
	const result = await inspectDocxReview(path);
	assert.equal(result.totalComments, 0);
	assert.deepEqual(result.comments, []);
	assert.equal(result.nextOffset, null);
});
