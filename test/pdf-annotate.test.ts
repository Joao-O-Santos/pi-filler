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
import { annotatePdf, matchPhrase, parseWords } from "../src/pdf-annotate.js";
import { runCommand } from "../src/process.js";

async function commandAvailable(name: string): Promise<boolean> {
	return runCommand(name, ["-v"]).then(
		() => true,
		() => false,
	);
}

async function makePdf(
	path: string,
	options: { rotate?: number; sign?: boolean } = {},
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
	await writeFile(path, await doc.save());
}

async function firstPageAnnotations(path: string): Promise<PDFDict[]> {
	const doc = await PDFDocument.load(await readFile(path));
	const annots = doc.getPage(0).node.Annots();
	const result: PDFDict[] = [];
	for (let index = 0; index < (annots?.size() ?? 0); index++) {
		const dict = doc.context.lookup(annots?.get(index), PDFDict);
		if (dict) result.push(dict);
	}
	return result;
}

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

test("adds a highlight comment to an unrotated PDF and keeps the input", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		const before = await readFile(input);

		const result = await annotatePdf(
			input,
			output,
			[
				{
					page: 1,
					text: "quarterly revenue",
					comment: "Check 10% “net” figure",
					author: "Reviewer",
				},
			],
			false,
		);
		assert.deepEqual(result, { output, dryRun: false, commentCount: 1, pages: [1] });
		assert.deepEqual(await readFile(input), before);

		const [annotation] = await firstPageAnnotations(output);
		assert.ok(annotation);
		assert.equal(annotation.get(PDFName.of("Subtype"))?.toString(), "/Highlight");
		const contents = annotation.get(PDFName.of("Contents"));
		assert.ok(contents instanceof PDFHexString);
		assert.equal(contents.decodeText(), "Check 10% “net” figure");
		const author = annotation.get(PDFName.of("T"));
		assert.ok(author instanceof PDFHexString);
		assert.equal(author.decodeText(), "Reviewer");
		const quads = annotation.get(PDFName.of("QuadPoints"));
		assert.ok(quads instanceof PDFArray);
		assert.equal(quads.size(), 16, "one quad (8 numbers) per highlighted word");
		assert.ok(annotation.get(PDFName.of("AP")));

		const text = await runCommand("pdftotext", ["-f", "1", "-l", "1", output, "-"]);
		assert.match(text.stdout.toString("utf8"), /quarterly revenue/);
		const leftovers = (await readdir(directory)).filter((name) => name.endsWith(".tmp"));
		assert.deepEqual(leftovers, []);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("dry runs validate without writing and reject unsupported comments", async (t) => {
	if (!(await commandAvailable("pdftotext"))) return t.skip("pdftotext is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdf-comment-dry-"));
	try {
		const input = join(directory, "input.pdf");
		const output = join(directory, "output.pdf");
		await makePdf(input);
		const dry = await annotatePdf(
			input,
			output,
			[{ page: 1, text: "fox jumps", comment: "note", author: "A" }],
			true,
		);
		assert.deepEqual(dry, { dryRun: true, commentCount: 1, pages: [1] });
		await assert.rejects(() => readFile(output), /ENOENT/);

		const comment = { page: 1, text: "quarterly revenue", comment: "note", author: "A" };
		await assert.rejects(
			() => annotatePdf(input, output, [comment, { ...comment, text: "revenue" }], true),
			/comment 2 \(page 1\): text matched 2 times/,
		);
		await assert.rejects(
			() => annotatePdf(input, output, [{ ...comment, page: 2 }], true),
			/page 2 exceeds 1 pages/,
		);
		await assert.rejects(
			() => annotatePdf(input, output, [{ ...comment, text: "not on the page" }], true),
			/comment 1 \(page 1\): text was not found on the page/,
		);

		const rotated = join(directory, "rotated.pdf");
		await makePdf(rotated, { rotate: 90 });
		await assert.rejects(
			() => annotatePdf(rotated, output, [comment], true),
			/rotated pages are not supported/,
		);

		const signed = join(directory, "signed.pdf");
		await makePdf(signed, { sign: true });
		await assert.rejects(
			() => annotatePdf(signed, output, [comment], true),
			/signature fields are not supported/,
		);
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("filler routes PDF patch through validated comment inputs", async () => {
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
			/requires pdf_comments/,
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
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});

test("schema and prompt guidance describe PDF comments", () => {
	const properties = fillerSchema.properties as unknown as Record<string, { maxItems?: number }>;
	assert.equal(properties.pdf_comments?.maxItems, 100);
	let guidelines: string[] = [];
	extension({
		registerTool(definition: { promptGuidelines?: string[] }) {
			guidelines = definition.promptGuidelines ?? [];
		},
	} as never);
	assert.ok(guidelines.some((line) => /PDF comments.*pdf_comments/.test(line)));
});
