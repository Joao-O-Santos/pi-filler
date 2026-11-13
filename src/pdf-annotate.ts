import { randomUUID } from "node:crypto";
import { readFile, rename, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import {
	PDFDict,
	PDFDocument,
	PDFHexString,
	PDFName,
	type PDFRef,
	PDFString,
} from "@cantoo/pdf-lib";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";

export interface PdfComment {
	page: number;
	text: string;
	comment: string;
	author: string;
}

export interface PdfCommentEdit {
	id: string;
	comment: string;
}

export interface PdfCommentRef {
	id: string;
	page: number;
}

export interface PdfCommentResult {
	output?: string;
	dryRun: boolean;
	commentCount: number;
	editCount: number;
	pages: number[];
	comments?: PdfCommentRef[];
}

export interface PdfOwnComment {
	id: string;
	page: number;
	author: string;
	comment: string;
	anchor: string;
	modified?: string;
}

// Annotations created by this tool carry this /NM prefix; only they can be
// edited or listed. The anchored phrase lives in a private /PiFillerAnchor key.
const NAME_PREFIX = "pi-filler:";
const ANCHOR = PDFName.of("PiFillerAnchor");

interface Word {
	text: string;
	xMin: number;
	yMin: number;
	xMax: number;
	yMax: number;
}

const SIG_FLAGS = PDFName.of("SigFlags");

const ENTITIES: Record<string, string> = {
	amp: "&",
	lt: "<",
	gt: ">",
	quot: '"',
	apos: "'",
};

function decodeEntities(text: string): string {
	return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, name: string) => {
		if (name.startsWith("#x")) return String.fromCodePoint(Number.parseInt(name.slice(2), 16));
		if (name.startsWith("#")) return String.fromCodePoint(Number(name.slice(1)));
		return ENTITIES[name.toLowerCase()] ?? match;
	});
}

// Parses word boxes from `pdftotext -bbox-layout` (top-left origin, points).
export function parseWords(html: string): { width: number; height: number; words: Word[] } {
	const page = html.match(/<page width="([\d.]+)" height="([\d.]+)"/);
	if (!page) throw new Error("pdftotext produced no page geometry");
	const words = [
		...html.matchAll(
			/<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<\/word>/g,
		),
	].map((match) => ({
		text: decodeEntities(match[5] ?? ""),
		xMin: Number(match[1]),
		yMin: Number(match[2]),
		xMax: Number(match[3]),
		yMax: Number(match[4]),
	}));
	return { width: Number(page[1]), height: Number(page[2]), words };
}

// Finds the single run of words whose text equals the whitespace-separated phrase.
export function matchPhrase(words: Word[], phrase: string): Word[] {
	const tokens = phrase.trim().split(/\s+/);
	const matches: Word[][] = [];
	for (let start = 0; start + tokens.length <= words.length; start++) {
		if (tokens.every((token, offset) => words[start + offset]?.text === token)) {
			matches.push(words.slice(start, start + tokens.length));
		}
	}
	if (matches.length !== 1) {
		throw new Error(
			matches.length === 0
				? "text was not found on the page"
				: `text matched ${matches.length} times on the page; use a more specific phrase`,
		);
	}
	return matches[0] ?? [];
}

async function pageWords(
	source: string,
	page: number,
	options: RunCommandOptions,
): Promise<{ width: number; height: number; words: Word[] }> {
	const result = await runCommand(
		"pdftotext",
		["-bbox-layout", "-f", String(page), "-l", String(page), source, "-"],
		options,
	);
	if (result.code !== 0) throw commandError(result, "pdftotext");
	return parseWords(result.stdout.toString("utf8"));
}

function pdfDate(date: Date): string {
	const pad = (value: number) => String(value).padStart(2, "0");
	return (
		`D:${date.getUTCFullYear()}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}` +
		`${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z`
	);
}

// Highlight appearance: one multiply-blended yellow quad per word.
function highlightAppearance(doc: PDFDocument, rect: number[], quads: number[][]): PDFRef {
	const ctx = doc.context;
	const graphicsState = ctx.obj({ Type: "ExtGState", BM: "Multiply" });
	const operators = quads
		.map(
			([x1, y1, x2, y2, x3, y3, x4, y4]) =>
				`${x1} ${y1} m ${x2} ${y2} l ${x4} ${y4} l ${x3} ${y3} l h f`,
		)
		.join("\n");
	const content = `/GS gs 1 1 0 rg\n${operators}\n`;
	return ctx.register(
		ctx.flateStream(content, {
			Type: "XObject",
			Subtype: "Form",
			BBox: rect,
			Resources: { ExtGState: { GS: ctx.register(graphicsState) } },
		}),
	);
}

function textOf(value: unknown): string | undefined {
	return value instanceof PDFHexString || value instanceof PDFString
		? value.decodeText()
		: undefined;
}

interface OwnAnnotation {
	id: string;
	page: number;
	dict: PDFDict;
}

// Lists this tool's annotations as (id, page, dictionary) without touching others.
function ownAnnotations(doc: PDFDocument): OwnAnnotation[] {
	const found: OwnAnnotation[] = [];
	const pages = doc.getPages();
	for (const [index, page] of pages.entries()) {
		const annots = page.node.Annots();
		for (let position = 0; position < (annots?.size() ?? 0); position++) {
			const dict = doc.context.lookup(annots?.get(position), PDFDict);
			const id = textOf(dict?.get(PDFName.of("NM")));
			if (dict && id?.startsWith(NAME_PREFIX)) {
				found.push({ id, page: index + 1, dict });
			}
		}
	}
	return found;
}

export async function listPdfComments(path: string): Promise<PdfOwnComment[]> {
	const doc = await PDFDocument.load(await readFile(resolve(path)), { updateMetadata: false });
	return ownAnnotations(doc).map(({ id, page, dict }) => ({
		id,
		page,
		author: textOf(dict.get(PDFName.of("T"))) ?? "",
		comment: textOf(dict.get(PDFName.of("Contents"))) ?? "",
		anchor: textOf(dict.get(ANCHOR)) ?? "",
		modified: textOf(dict.get(PDFName.of("M"))),
	}));
}

function assertSupportedDocument(doc: PDFDocument, pageCount: number, comments: PdfComment[]) {
	const acroForm = doc.catalog.lookup(PDFName.of("AcroForm"));
	if (acroForm instanceof PDFDict && acroForm.get(SIG_FLAGS) !== undefined) {
		throw new Error("PDFs with signature fields are not supported");
	}
	for (const [index, comment] of comments.entries()) {
		if (comment.page > pageCount) {
			throw new Error(`comment ${index + 1}: page ${comment.page} exceeds ${pageCount} pages`);
		}
		const rotation = doc.getPage(comment.page - 1).getRotation().angle;
		if (rotation !== 0) {
			throw new Error(`comment ${index + 1}: rotated pages are not supported`);
		}
	}
}

export interface PdfCommentChanges {
	comments?: PdfComment[];
	edits?: PdfCommentEdit[];
}

// Adds Highlight annotations and edits this tool's own comments in a copy of
// `path`. Each new comment must match exactly one phrase occurrence on its page.
// Output is published atomically; dry runs validate without writing or
// returning new ids.
export async function annotatePdf(
	path: string,
	output: string | undefined,
	changes: PdfCommentChanges,
	dryRun: boolean,
	options: RunCommandOptions = {},
): Promise<PdfCommentResult> {
	const comments = changes.comments ?? [];
	const edits = changes.edits ?? [];
	const source = resolve(path);
	const bytes = await readFile(source);
	const doc = await PDFDocument.load(bytes, { updateMetadata: false });
	assertSupportedDocument(doc, doc.getPageCount(), comments);

	const owned = new Map(ownAnnotations(doc).map((entry) => [entry.id, entry]));
	const seen = new Set<string>();
	for (const edit of edits) {
		if (seen.has(edit.id)) throw new Error(`comment id ${edit.id} is edited more than once`);
		seen.add(edit.id);
		if (!owned.has(edit.id)) throw new Error(`comment id ${edit.id} was not found`);
	}

	const date = pdfDate(new Date());
	for (const edit of edits) {
		const target = owned.get(edit.id)?.dict;
		target?.set(PDFName.of("Contents"), PDFHexString.fromText(edit.comment));
		target?.set(PDFName.of("M"), PDFString.of(date));
	}

	const wordsByPage = new Map<number, Word[]>();
	const pages = new Set<number>();
	const refs: PdfCommentRef[] = [];
	for (const [index, comment] of comments.entries()) {
		const label = `comment ${index + 1} (page ${comment.page})`;
		const page = doc.getPage(comment.page - 1);
		const box = page.getCropBox();
		if (!wordsByPage.has(comment.page)) {
			const extracted = await pageWords(source, comment.page, options);
			wordsByPage.set(comment.page, extracted.words);
		}
		let matched: Word[];
		try {
			matched = matchPhrase(wordsByPage.get(comment.page) ?? [], comment.text);
		} catch (error) {
			throw new Error(`${label}: ${(error as Error).message}`);
		}

		const quads = matched.map((word) => {
			const left = box.x + word.xMin;
			const right = box.x + word.xMax;
			const top = box.y + box.height - word.yMin;
			const bottom = box.y + box.height - word.yMax;
			return [left, top, right, top, left, bottom, right, bottom];
		});
		const xs = quads.flatMap((quad) => [quad[0] ?? 0, quad[2] ?? 0]);
		const ys = quads.flatMap((quad) => [quad[1] ?? 0, quad[5] ?? 0]);
		const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

		const ctx = doc.context;
		const id = `${NAME_PREFIX}${randomUUID()}`;
		const annotation = ctx.register(
			ctx.obj({
				Type: "Annot",
				Subtype: "Highlight",
				Rect: rect,
				QuadPoints: quads.flat(),
				C: [1, 1, 0],
				CA: 1,
				F: 4,
				NM: PDFHexString.fromText(id),
				T: PDFHexString.fromText(comment.author),
				Contents: PDFHexString.fromText(comment.comment),
				PiFillerAnchor: PDFHexString.fromText(comment.text),
				M: PDFString.of(date),
				AP: { N: highlightAppearance(doc, rect, quads) },
			}),
		);
		page.node.addAnnot(annotation);
		pages.add(comment.page);
		refs.push({ id, page: comment.page });
	}

	const result: PdfCommentResult = {
		dryRun,
		commentCount: comments.length,
		editCount: edits.length,
		pages: [...pages].sort((a, b) => a - b),
	};
	if (dryRun) return result;
	if (!output) throw new Error("PDF comment output is required");

	const target = resolve(output);
	const temporary = join(dirname(target), `.${target.split(/[\\/]/).pop()}.${randomUUID()}.tmp`);
	try {
		await writeFile(temporary, await doc.save());
		await rename(temporary, target);
	} finally {
		await rm(temporary, { force: true });
	}
	return { ...result, output: target, comments: refs };
}
