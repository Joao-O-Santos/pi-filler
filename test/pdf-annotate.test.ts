import assert from "node:assert/strict";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import {
	degrees,
	PDFArray,
	PDFDict,
	PDFDocument,
	PDFHexString,
	PDFName,
	type PDFRef,
	StandardFonts,
} from "@cantoo/pdf-lib";
import extension, { executeFiller, fillerSchema } from "../src/index.js";
import { annotatePdf, listPdfComments, matchPhrase, parseWords } from "../src/pdf-annotate.js";
import { runCommand } from "../src/process.js";

async function commandAvailable(name: string): Promise<boolean> {
	return runCommand(name, ["-v"]).then(
		() => true,
		() => false,
	);
}

async function makePdf(
	path: string,
	options: { rotate?: number; sign?: boolean; foreignHighlight?: boolean } = {},
): Promise<void> {
	const doc = await PDFDocument.create();
	const font = await doc.embedFont(StandardFonts.Helvetica);
	const page = doc.addPage([612, 792]);
	const lines = [
		"The quick brown fox jumps over the lazy dog",
		"Second line with the quarterly revenue figure",
		"Quarterly revenue appears again",
	];
	for (const [index, line] of lines.entries()) {
		page.drawText(line, { x: 72, y: 700 - index * 30, size: 14, font });
	}
	if (options.rotate !== undefined) page.setRotation(degrees(options.rotate));
	if (options.sign) {
		doc.catalog.set(PDFName.of("AcroForm"), doc.context.obj({ SigFlags: 3, Fields: [] }));
	}
	if (options.foreignHighlight) {
		const annotation = doc.context.register(
			doc.context.obj({
				Type: "Annot",
				Subtype: "Highlight",
				Rect: [72, 690, 120, 705],
				QuadPoints: [72, 705, 120, 705, 72, 690, 120, 690],
				Contents: PDFHexString.fromText("someone else's note"),
				NM: PDFHexString.fromText("reviewer-1"),
			}),
		);
		page.node.addAnnot(annotation);
	}
	await writeFile(path, await doc.save());
}

interface TestPdfAnnotation {
	subtype?: string;
	name?: string;
	comment?: string;
	richText?: string;
	indirect?: boolean;
	richStream?: boolean;
	unsupportedRich?: boolean;
	appearance?: boolean;
	replyTo?: number;
	direct?: boolean;
	page?: number;
}

async function addTestAnnotations(path: string, entries: TestPdfAnnotation[]): Promise<void> {
	const doc = await PDFDocument.load(await readFile(path));
	if (entries.some((entry) => (entry.page ?? 1) > 1)) doc.addPage();
	const refs: PDFRef[] = [];
	for (const entry of entries) {
		const page = doc.getPage((entry.page ?? 1) - 1);
		const indirectString = (value: string) =>
			entry.indirect
				? doc.context.register(PDFHexString.fromText(value))
				: PDFHexString.fromText(value);
		const rich = entry.unsupportedRich
			? doc.context.register(doc.context.obj({ Unsupported: true }))
			: entry.richText === undefined
				? undefined
				: entry.richStream
					? doc.context.register(doc.context.flateStream(entry.richText))
					: indirectString(entry.richText);
		const annotation = doc.context.obj({
			Type: "Annot",
			Subtype: PDFName.of(entry.subtype ?? "Highlight"),
			Rect: [72, 690, 120, 705],
			...(entry.name === undefined ? {} : { NM: indirectString(entry.name) }),
			...(entry.indirect ? { T: indirectString("Indirect author") } : {}),
			...(entry.comment === undefined ? {} : { Contents: indirectString(entry.comment) }),
			...(rich === undefined ? {} : { RC: rich }),
			...(entry.appearance
				? { AP: { N: doc.context.register(doc.context.flateStream("old visible text")) } }
				: {}),
			...(entry.replyTo === undefined ? {} : { IRT: refs[entry.replyTo], RT: PDFName.of("R") }),
		});
		const reference = doc.context.register(annotation);
		refs.push(reference);
		if (entry.direct) page.node.Annots()?.push(annotation);
		else page.node.addAnnot(reference);
	}
	await writeFile(path, await doc.save());
}

async function pageOneAnnotations(path: string): Promise<PDFDict[]> {
	const doc = await PDFDocument.load(await readFile(path));
	const annots = doc.getPage(0).node.Annots();
	const result: PDFDict[] = [];
	for (let index = 0; index < (annots?.size() ?? 0); index++) {
		const dict = doc.context.lookup(annots?.get(index), PDFDict);
		if (dict) result.push(dict);
	}
	return result;
}

const quarterly = {
	page: 1,
	text: "quarterly revenue",
	comment: "Check figure",
	author: "Reviewer",
};

test("parses pdftotext word boxes and decodes entities", () => {
	const html = `<page width="612.000000" height="792.000000">
<word xMin="1.5" yMin="2.5" xMax="3.5" yMax="4.5">Tom&amp;Jerry</word>
<word xMin="5" yMin="6" xMax="7" yMax="8">&#8212;&#x41;</word>
</page>`;
	const parsed = parseWords(html);
	assert.equal(parsed.width, 612);
	assert.equal(parsed.height, 792);
	assert.deepEqual(
		parsed.words.map((word) => word.text),
		["Tom&Jerry", "—A"],
	);
	assert.throws(() => parseWords("<doc/>"), /no page geometry/);
});

test("matches one exact phrase and rejects missing or repeated phrases", () => {
	const words = ["a", "quarterly", "revenue", "b", "quarterly", "revenue."].map((text, index) => ({
		text,
		xMin: index,
		yMin: 0,
		xMax: index + 1,
		yMax: 1,
	}));
	assert.deepEqual(
		matchPhrase(words, "revenue.").map((word) => word.text),
		["revenue."],
	);
	assert.throws(() => matchPhrase(words, "missing"), /not found/);
	assert.throws(() => matchPhrase(words, "quarterly"), /matched 2 times/);
});

test("adds owned highlight comments, lists them, and edits them in place", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-"));
	try {
		const input = join(directory, "input.pdf");
		const first = join(directory, "first.pdf");
		const second = join(directory, "second.pdf");
		await makePdf(input, { foreignHighlight: true });
		const before = await readFile(input);

		const added = await annotatePdf(
			input,
			first,
			{ comments: [{ ...quarterly, comment: "Check 10% “net” figure" }] },
			false,
		);
		assert.equal(added.output, first);
		assert.equal(added.commentCount, 1);
		assert.deepEqual(added.pages, [1]);
		assert.equal(added.comments?.length, 1);
		const id = added.comments?.[0]?.id ?? "";
		assert.match(id, /^pi-filler:[0-9a-f-]{36}$/);
		assert.deepEqual(await readFile(input), before);

		const listed = await listPdfComments(first);
		const foreignId = listed.find((entry) => entry.comment === "someone else's note")?.id;
		assert.ok(foreignId?.startsWith("pdf-comment:v1:"));
		assert.deepEqual(
			listed.map((entry) => ({
				id: entry.id,
				page: entry.page,
				author: entry.author,
				comment: entry.comment,
				anchor: entry.anchor,
			})),
			[
				{
					id: foreignId,
					page: 1,
					author: "",
					comment: "someone else's note",
					anchor: "",
				},
				{
					id,
					page: 1,
					author: "Reviewer",
					comment: "Check 10% “net” figure",
					anchor: "quarterly revenue",
				},
			],
			"foreign comments are listed too",
		);
		assert.ok(listed.find((entry) => entry.id === id)?.modified?.startsWith("D:"));

		const [annotation] = (await pageOneAnnotations(first)).filter(
			(dict) =>
				dict.get(PDFName.of("NM"))?.toString() !== undefined &&
				dict.get(PDFName.of("Subtype"))?.toString() === "/Highlight" &&
				dict.get(PDFName.of("PiFillerAnchor")),
		);
		assert.ok(annotation);
		const quads = annotation.get(PDFName.of("QuadPoints"));
		assert.ok(quads instanceof PDFArray);
		assert.equal(quads.size(), 16, "one quad (8 numbers) per highlighted word");
		assert.ok(annotation.get(PDFName.of("AP")));
		const rect = annotation.get(PDFName.of("Rect"))?.toString();

		const edited = await annotatePdf(
			first,
			second,
			{ edits: [{ id, comment: "Revised: figure now audited" }] },
			false,
		);
		assert.equal(edited.editCount, 1);
		assert.deepEqual(edited.comments, [], "edits do not return new ids");
		const relisted = await listPdfComments(second);
		const relistedOwn = relisted.find((entry) => entry.id === id);
		assert.ok(relistedOwn, "ids are stable across edits");
		assert.equal(relistedOwn.comment, "Revised: figure now audited");
		assert.equal(relistedOwn.author, "Reviewer");
		assert.equal(relistedOwn.anchor, "quarterly revenue");

		const [editedAnnotation] = (await pageOneAnnotations(second)).filter((dict) =>
			dict.get(PDFName.of("PiFillerAnchor")),
		);
		assert.equal(editedAnnotation?.get(PDFName.of("Rect"))?.toString(), rect);

		const text = await runCommand("pdftotext", ["-f", "1", "-l", "1", second, "-"]);
		assert.match(text.stdout.toString("utf8"), /quarterly revenue/);
		const leftovers = (await readdir(directory)).filter((name) => name.endsWith(".tmp"));
		assert.deepEqual(leftovers, []);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("lists and edits an existing third-party highlight comment", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-foreign-comment-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input, { foreignHighlight: true });
		const before = await readFile(input);
		const comments = await listPdfComments(input);
		assert.equal(comments.length, 1);
		assert.equal(comments[0]?.comment, "someone else's note");
		const id = comments[0]?.id;
		assert.ok(id);
		await annotatePdf(input, output, { edits: [{ id, comment: "Revised externally" }] }, false);
		const edited = await listPdfComments(output);
		assert.equal(edited[0]?.comment, "Revised externally");
		assert.deepEqual(await readFile(input), before);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("lists foreign annotations with duplicate and missing names and preserves IDs after edits", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-foreign-comments-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		await addTestAnnotations(input, [
			{ name: "duplicate", comment: "first" },
			{ name: "duplicate", comment: "second" },
			{ comment: "nameless" },
			{ comment: "direct dictionary", direct: true },
			{ name: "rich", richText: "<body><p>Rich &amp; formatted</p></body>" },
			{ name: "reply-parent", comment: "parent" },
			{ name: "reply-child", comment: "reply", replyTo: 5 },
			{ subtype: "Popup", comment: "popup duplicate" },
			{ subtype: "Widget", comment: "field tooltip" },
			{ subtype: "Text", comment: "sticky note" },
			{ subtype: "FreeText", comment: "free text note" },
			{ name: "second-page", comment: "second page note", page: 2 },
		]);
		const listed = await listPdfComments(input);
		assert.deepEqual(
			listed.map((comment) => comment.comment),
			[
				"first",
				"second",
				"nameless",
				"direct dictionary",
				"Rich & formatted",
				"parent",
				"reply",
				"sticky note",
				"free text note",
				"second page note",
			],
		);
		assert.equal(new Set(listed.map((comment) => comment.id)).size, listed.length);
		assert.ok(listed[0]?.id.startsWith("pdf-comment:v1:"));
		assert.ok(listed[1]?.id.startsWith("pdf-comment:v1:"));
		assert.ok(listed[2]?.id.startsWith("pdf-comment:v1:"));
		assert.ok(listed[3]?.id.startsWith("pdf-comment:v1:"));
		assert.ok(listed[4]?.id.startsWith("pdf-comment:v1:"));
		assert.equal(listed[6]?.inReplyTo, listed[5]?.id);
		assert.equal(listed[7]?.subtype, "/Text");
		assert.equal(listed[8]?.subtype, "/FreeText");
		assert.equal(listed[9]?.page, 2);
		assert.ok(listed[9]?.id.startsWith("pdf-comment:v1:"));

		const duplicateId = listed[0]?.id ?? "";
		const richId = listed[4]?.id ?? "";
		const directId = listed[3]?.id ?? "";
		await annotatePdf(
			input,
			output,
			{
				edits: [
					{ id: duplicateId, comment: "Only the first duplicate changed" },
					{ id: richId, comment: "Plain revised text" },
					{ id: directId, comment: "Direct entry revised" },
				],
			},
			false,
		);
		const edited = await listPdfComments(output);
		assert.equal(edited[0]?.id, duplicateId);
		assert.equal(edited[0]?.comment, "Only the first duplicate changed");
		assert.equal(edited[1]?.comment, "second");
		assert.equal(edited[3]?.id, directId);
		assert.equal(edited[3]?.comment, "Direct entry revised");
		assert.equal(edited[4]?.id, richId);
		assert.equal(edited[4]?.comment, "Plain revised text");
		const dicts = await pageOneAnnotations(output);
		const richAnnotation = dicts.find(
			(dict) => dict.lookupMaybe(PDFName.of("NM"), PDFHexString)?.decodeText() === richId,
		);
		assert.ok(richAnnotation);
		assert.equal(
			richAnnotation.get(PDFName.of("RC")),
			undefined,
			"stale rich text is removed on edit",
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("rejects colliding effective ids instead of editing either annotation", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-comment-id-collision-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		await addTestAnnotations(input, [{ comment: "first" }, { comment: "second" }]);
		const firstId = (await listPdfComments(input))[0]?.id;
		assert.ok(firstId);
		const doc = await PDFDocument.load(await readFile(input));
		const annots = doc.getPage(0).node.Annots();
		const second = doc.context.lookup(annots?.get(1), PDFDict);
		assert.ok(second);
		second.set(PDFName.of("NM"), PDFHexString.fromText(firstId));
		await writeFile(input, await doc.save());
		await assert.rejects(() => listPdfComments(input), /duplicate effective ids/);
		await assert.rejects(
			() =>
				annotatePdf(input, output, { edits: [{ id: firstId, comment: "wrong target" }] }, false),
			/duplicate effective ids/,
		);
		await assert.rejects(() => readFile(output), /ENOENT/);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("resolves indirect strings and rich-text streams and removes stale FreeText appearance", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-indirect-comment-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		await addTestAnnotations(input, [
			{ name: "pi-filler:indirect-name", comment: "indirect body", indirect: true },
			{ richText: "<body><p>Stream &amp; body</p></body>", richStream: true },
			{ subtype: "FreeText", comment: "old visible body", appearance: true },
		]);
		const listed = await listPdfComments(input);
		assert.equal(listed[0]?.id, "pi-filler:indirect-name");
		assert.equal(listed[0]?.author, "Indirect author");
		assert.equal(listed[0]?.comment, "indirect body");
		assert.equal(listed[1]?.comment, "Stream & body");
		const freeTextId = listed[2]?.id;
		assert.ok(freeTextId);
		await annotatePdf(
			input,
			output,
			{ edits: [{ id: freeTextId, comment: "new visible body" }] },
			false,
		);
		const annotations = await listPdfComments(output);
		assert.equal(annotations[2]?.comment, "new visible body");
		const dicts = await pageOneAnnotations(output);
		const freeText = dicts.find(
			(dict) => dict.get(PDFName.of("Subtype"))?.toString() === "/FreeText",
		);
		assert.ok(freeText);
		assert.equal(
			freeText.get(PDFName.of("AP")),
			undefined,
			"old appearance cannot display stale text",
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("rejects unsupported rich-text stream objects", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-unsupported-rich-comment-"));
	try {
		const input = join(directory, "input.pdf");
		await makePdf(input);
		await addTestAnnotations(input, [{ unsupportedRich: true }]);
		await assert.rejects(
			() => listPdfComments(input),
			/unsupported PDF rich-text comment representation/,
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("rejects undecodable rich-text streams without publishing an edit", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-bad-rich-stream-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		const doc = await PDFDocument.load(await readFile(input));
		const stream = doc.context.register(
			doc.context.stream("not compressed", { Filter: "FlateDecode" }),
		);
		const annotation = doc.context.register(
			doc.context.obj({
				Type: "Annot",
				Subtype: "Text",
				Rect: [1, 1, 10, 10],
				RC: stream,
			}),
		);
		doc.getPage(0).node.addAnnot(annotation);
		await writeFile(input, await doc.save());
		await assert.rejects(() => listPdfComments(input), /unsupported PDF rich-text comment stream/);
		await assert.rejects(
			() => annotatePdf(input, output, { edits: [{ id: "unknown", comment: "x" }] }, false),
			/unsupported PDF rich-text comment stream/,
		);
		await assert.rejects(() => readFile(output), /ENOENT/);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("stale content-bound IDs fail closed before publication", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-stale-comment-"));
	try {
		const input = join(directory, "input.pdf");
		const changed = join(directory, "changed.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		await addTestAnnotations(input, [{ comment: "original" }]);
		const id = (await listPdfComments(input))[0]?.id ?? "";
		const doc = await PDFDocument.load(await readFile(input));
		const annots = doc.getPage(0).node.Annots();
		const annotation = doc.context.lookup(annots?.get((annots?.size() ?? 1) - 1), PDFDict);
		assert.ok(annotation);
		annotation.set(PDFName.of("Contents"), PDFHexString.fromText("changed elsewhere"));
		await writeFile(changed, await doc.save());
		await assert.rejects(
			() => annotatePdf(changed, output, { edits: [{ id, comment: "stale edit" }] }, false),
			/comment id .* was not found/,
		);
		await assert.rejects(() => readFile(output), /ENOENT/);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("adds and edits in one atomic call; dry runs never write", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-mixed-"));
	try {
		const input = join(directory, "input.pdf");
		const first = join(directory, "first.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		const base = await annotatePdf(input, first, { comments: [quarterly] }, false);
		const id = base.comments?.[0]?.id ?? "";

		const dry = await annotatePdf(
			first,
			output,
			{
				edits: [{ id, comment: "changed" }],
				comments: [{ page: 1, text: "fox jumps", comment: "new", author: "A" }],
			},
			true,
		);
		assert.deepEqual(dry, { dryRun: true, commentCount: 1, editCount: 1, pages: [1] });
		await assert.rejects(() => readFile(output), /ENOENT/);

		const mixed = await annotatePdf(
			first,
			output,
			{
				edits: [{ id, comment: "changed" }],
				comments: [{ page: 1, text: "fox jumps", comment: "new", author: "A" }],
			},
			false,
		);
		assert.equal(mixed.comments?.length, 1);
		const listed = await listPdfComments(output);
		assert.deepEqual(listed.map((entry) => entry.comment).sort(), ["changed", "new"]);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("rejects unknown and duplicate IDs while allowing foreign comment edits", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-ids-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input, { foreignHighlight: true });
		const added = await annotatePdf(
			input,
			join(directory, "owned.pdf"),
			{ comments: [quarterly] },
			false,
		);
		const id = added.comments?.[0]?.id ?? "";
		const owned = join(directory, "owned.pdf");
		const foreignId = (await listPdfComments(owned)).find(
			(comment) => comment.comment === "someone else's note",
		)?.id;
		assert.ok(foreignId);

		await assert.rejects(
			() =>
				annotatePdf(owned, output, { edits: [{ id: "pi-filler:missing", comment: "x" }] }, true),
			/comment id pi-filler:missing was not found/,
		);
		const foreignOutput = join(directory, "foreign-output.pdf");
		await annotatePdf(owned, foreignOutput, { edits: [{ id: foreignId, comment: "x" }] }, false);
		assert.equal(
			(await listPdfComments(foreignOutput)).find((comment) => comment.id === foreignId)?.comment,
			"x",
		);
		await assert.rejects(
			() =>
				annotatePdf(
					owned,
					output,
					{
						edits: [
							{ id, comment: "x" },
							{ id, comment: "y" },
						],
					},
					true,
				),
			/edited more than once/,
		);
		await assert.rejects(() => readFile(output), /ENOENT/);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("rejects unsupported comment targets without writing", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-dry-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		const dry = await annotatePdf(
			input,
			output,
			{ comments: [{ page: 1, text: "fox jumps", comment: "note", author: "A" }] },
			true,
		);
		assert.deepEqual(dry, { dryRun: true, commentCount: 1, editCount: 0, pages: [1] });
		await assert.rejects(() => readFile(output), /ENOENT/);

		const comment = { page: 1, text: "quarterly revenue", comment: "note", author: "A" };
		await assert.rejects(
			() =>
				annotatePdf(input, output, { comments: [comment, { ...comment, text: "revenue" }] }, true),
			/comment 2 \(page 1\): text matched 2 times/,
		);
		await assert.rejects(
			() => annotatePdf(input, output, { comments: [{ ...comment, page: 2 }] }, true),
			/page 2 exceeds 1 pages/,
		);
		await assert.rejects(
			() =>
				annotatePdf(input, output, { comments: [{ ...comment, text: "not on the page" }] }, true),
			/comment 1 \(page 1\): text was not found on the page/,
		);

		const rotated = join(directory, "rotated.pdf");
		await makePdf(rotated, { rotate: 90 });
		await assert.rejects(
			() => annotatePdf(rotated, output, { comments: [comment] }, true),
			/rotated pages are not supported/,
		);

		const signed = join(directory, "signed.pdf");
		await makePdf(signed, { sign: true });
		await assert.rejects(
			() => annotatePdf(signed, output, { comments: [comment] }, true),
			/signature fields are not supported/,
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("filler routes PDF comment patches and comment listing", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-route-"));
	try {
		const pdf = join(directory, "input.pdf");
		await makePdf(pdf);
		const comments = [{ page: 1, text: "fox", comment: "note", author: "A" }];
		const base = { format: "pdf" as const, action: "patch" as const, path: pdf };
		const out = join(directory, "o.pdf");
		await assert.rejects(
			() => executeFiller({ ...base, pdf_comments: comments }, directory),
			/PDF patch requires output/,
		);
		await assert.rejects(
			() => executeFiller({ ...base, output: pdf, pdf_comments: comments }, directory),
			/must differ from input/,
		);
		await assert.rejects(
			() => executeFiller({ ...base, output: out }, directory),
			/requires pdf_comments or pdf_comment_edits/,
		);
		await assert.rejects(
			() =>
				executeFiller({ ...base, output: out, pdf_comments: comments, view: "text" }, directory),
			/does not accept view/,
		);
		await assert.rejects(
			() => executeFiller({ ...base, output: out, pdf_comments: comments, query: "x" }, directory),
			/does not accept query/,
		);
		const result = await executeFiller(
			{ ...base, output: out, pdf_comments: comments, dry_run: true },
			directory,
		);
		assert.match(result.text, /"dryRun": true/);
		assert.doesNotMatch(result.text, /fox|note/);

		const added = await executeFiller({ ...base, output: out, pdf_comments: comments }, directory);
		assert.doesNotMatch(added.text, /fox|note/);
		const id = (added.details.comments as { id: string }[])[0]?.id ?? "";
		const edited = join(directory, "edited.pdf");
		await executeFiller(
			{
				...base,
				path: out,
				output: edited,
				pdf_comment_edits: [{ id, comment: "revised" }],
			},
			directory,
		);

		const listing = await executeFiller(
			{ format: "pdf", action: "read", view: "comments", path: edited },
			directory,
		);
		assert.deepEqual(
			JSON.parse(listing.text).map((entry: { comment: string }) => entry.comment),
			["revised"],
		);
		await assert.rejects(
			() =>
				executeFiller(
					{ format: "pdf", action: "read", view: "comments", path: edited, query: "x" },
					directory,
				),
			/does not accept query/,
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("schema and prompt guidance describe PDF comment creation, edits, and listing", () => {
	const properties = fillerSchema.properties as unknown as Record<
		string,
		{ maxItems?: number; description?: string }
	>;
	assert.equal(properties.pdf_comments?.maxItems, 100);
	assert.equal(properties.pdf_comment_edits?.maxItems, 100);
	assert.match(properties.pdf_comment_edits?.description ?? "", /other PDF applications/);
	let guidelines: string[] = [];
	extension({
		registerTool(definition: { promptGuidelines?: string[] }) {
			guidelines = definition.promptGuidelines ?? [];
		},
	} as never);
	assert.ok(guidelines.some((line) => /PDF comments.*pdf_comments/.test(line)));
	assert.ok(
		guidelines.some(
			(line) =>
				/pdf_comment_edits/.test(line) &&
				/view comments/.test(line) &&
				/other PDF applications/.test(line),
		),
	);
	assert.ok(guidelines.some((line) => /FreeText edits remove the cached appearance/.test(line)));
});
