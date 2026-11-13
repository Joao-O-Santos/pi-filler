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
					id,
					page: 1,
					author: "Reviewer",
					comment: "Check 10% “net” figure",
					anchor: "quarterly revenue",
				},
			],
			"foreign annotations are never listed",
		);
		assert.ok(listed[0]?.modified?.startsWith("D:"));

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
		assert.equal(relisted.length, 1);
		assert.equal(relisted[0]?.id, id, "ids are stable across edits");
		assert.equal(relisted[0]?.comment, "Revised: figure now audited");
		assert.equal(relisted[0]?.author, "Reviewer");
		assert.equal(relisted[0]?.anchor, "quarterly revenue");

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

test("rejects unknown, duplicate, and foreign comment ids", async (t) => {
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

		await assert.rejects(
			() =>
				annotatePdf(owned, output, { edits: [{ id: "pi-filler:missing", comment: "x" }] }, true),
			/comment id pi-filler:missing was not found/,
		);
		await assert.rejects(
			() => annotatePdf(owned, output, { edits: [{ id: "reviewer-1", comment: "x" }] }, true),
			/comment id reviewer-1 was not found/,
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
	const properties = fillerSchema.properties as unknown as Record<string, { maxItems?: number }>;
	assert.equal(properties.pdf_comments?.maxItems, 100);
	assert.equal(properties.pdf_comment_edits?.maxItems, 100);
	let guidelines: string[] = [];
	extension({
		registerTool(definition: { promptGuidelines?: string[] }) {
			guidelines = definition.promptGuidelines ?? [];
		},
	} as never);
	assert.ok(guidelines.some((line) => /PDF comments.*pdf_comments/.test(line)));
	assert.ok(
		guidelines.some((line) => /pdf_comment_edits/.test(line) && /view=comments/.test(line)),
	);
});
