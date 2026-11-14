import { createHash, randomUUID } from "node:crypto";
import { readFile, rename, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import {
	decodePDFRawStream,
	PDFDict,
	PDFDocument,
	PDFHexString,
	PDFName,
	PDFRawStream,
	PDFRef,
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

export interface PdfCommentRecord {
	id: string;
	page: number;
	subtype: string;
	author: string;
	comment: string;
	anchor: string;
	inReplyTo?: string;
	modified?: string;
}

// Preserve older pi-filler IDs. Foreign comments without a unique /NM get a
// content-bound locator that is promoted to /NM when edited.
const NAME_PREFIX = "pi-filler:";
const LOCATOR_PREFIX = "pdf-comment:v1:";
const ANCHOR = PDFName.of("PiFillerAnchor");
const CONTENTS = PDFName.of("Contents");
const RICH_CONTENTS = PDFName.of("RC");
const ANNOTATION_NAME = PDFName.of("NM");
const REPLY_TO = PDFName.of("IRT");
const EXCLUDED_SUBTYPES = new Set(["/Popup", "/Widget"]);

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
	nbsp: "\u00a0",
	ndash: "–",
	mdash: "—",
	lsquo: "‘",
	rsquo: "’",
	ldquo: "“",
	rdquo: "”",
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

function textOf(doc: PDFDocument, value: unknown): string | undefined {
	const resolved = value instanceof PDFRef ? doc.context.lookup(value) : value;
	return resolved instanceof PDFString || resolved instanceof PDFHexString
		? resolved.decodeText()
		: undefined;
}

interface PdfAnnotation {
	id: string;
	page: number;
	position: number;
	dict: PDFDict;
	subtype: string;
	rawName?: string;
	locator: boolean;
	ref?: PDFRef;
}

function annotationText(doc: PDFDocument, dict: PDFDict): string | undefined {
	const contents = textOf(doc, dict.get(CONTENTS));
	if (contents !== undefined) return contents;
	const richValue = dict.get(RICH_CONTENTS);
	let rich = textOf(doc, richValue);
	if (rich === undefined && richValue !== undefined) {
		const object = richValue instanceof PDFRef ? doc.context.lookup(richValue) : richValue;
		if (!(object instanceof PDFRawStream)) {
			throw new Error("unsupported PDF rich-text comment representation");
		}
		try {
			rich = new TextDecoder().decode(decodePDFRawStream(object).decode());
		} catch {
			throw new Error("unsupported PDF rich-text comment stream");
		}
	}
	if (rich === undefined) return undefined;
	return decodeEntities(
		rich
			.replace(/<\/(?:p|div|li)>/gi, "\n")
			.replace(/<br\s*\/?\s*>/gi, "\n")
			.replace(/<[^>]*>/g, ""),
	).trim();
}

function locatorId(doc: PDFDocument, dict: PDFDict, page: number, position: number): string {
	const fields = dict
		.entries()
		.filter(([key]) => ![PDFName.of("M"), PDFName.of("AP")].includes(key))
		.sort(([left], [right]) => left.toString().localeCompare(right.toString()))
		.map(([key, value]) => `${key.toString()}=${value.toString()}`)
		.join("\n");
	const comment = annotationText(doc, dict) ?? "";
	const author = textOf(doc, dict.get(PDFName.of("T"))) ?? "";
	const name = textOf(doc, dict.get(ANNOTATION_NAME)) ?? "";
	const hash = createHash("sha256")
		.update(`${fields}\n${comment}\n${author}\n${name}`)
		.digest("hex");
	return `${LOCATOR_PREFIX}${page}:${position}:${hash}`;
}

function annotations(doc: PDFDocument): PdfAnnotation[] {
	const raw: Array<Omit<PdfAnnotation, "id" | "rawName" | "locator"> & { rawName?: string }> = [];
	for (const [pageIndex, page] of doc.getPages().entries()) {
		const annots = page.node.Annots();
		for (let position = 0; position < (annots?.size() ?? 0); position++) {
			const object = annots?.get(position);
			const dict = doc.context.lookup(object, PDFDict);
			if (!dict) continue;
			const subtype = dict.get(PDFName.of("Subtype"))?.toString() ?? "";
			if (EXCLUDED_SUBTYPES.has(subtype) || annotationText(doc, dict) === undefined) continue;
			raw.push({
				page: pageIndex + 1,
				position,
				dict,
				subtype,
				...(textOf(doc, dict.get(ANNOTATION_NAME)) === undefined
					? {}
					: { rawName: textOf(doc, dict.get(ANNOTATION_NAME)) }),
				...(object instanceof PDFRef ? { ref: object } : {}),
			});
		}
	}
	const nameCounts = new Map<string, number>();
	for (const item of raw) {
		if (item.rawName !== undefined)
			nameCounts.set(item.rawName, (nameCounts.get(item.rawName) ?? 0) + 1);
	}
	const entries = raw.map((item) => {
		const persistentName =
			item.rawName?.startsWith(NAME_PREFIX) === true ||
			/^pdf-comment:v1:\d+:\d+:[a-f0-9]{64}$/.test(item.rawName ?? "");
		const useName =
			item.rawName !== undefined && nameCounts.get(item.rawName) === 1 && persistentName;
		const id = useName
			? (item.rawName as string)
			: locatorId(doc, item.dict, item.page, item.position);
		return { ...item, id, locator: !useName };
	});
	const ids = new Set<string>();
	for (const entry of entries) {
		if (ids.has(entry.id)) throw new Error("PDF comments have duplicate effective ids");
		ids.add(entry.id);
	}
	return entries;
}

export async function listPdfComments(path: string): Promise<PdfCommentRecord[]> {
	const doc = await PDFDocument.load(await readFile(resolve(path)), { updateMetadata: false });
	const entries = annotations(doc);
	const byRef = new Map(
		entries.flatMap((entry) => (entry.ref ? [[entry.ref.toString(), entry.id]] : [])),
	);
	return entries.map(({ id, page, dict, subtype }) => {
		const reply = dict.get(REPLY_TO);
		const inReplyTo = reply instanceof PDFRef ? byRef.get(reply.toString()) : undefined;
		return {
			id,
			page,
			subtype,
			author: textOf(doc, dict.get(PDFName.of("T"))) ?? "",
			comment: annotationText(doc, dict) ?? "",
			anchor: textOf(doc, dict.get(ANCHOR)) ?? "",
			...(inReplyTo === undefined ? {} : { inReplyTo }),
			modified: textOf(doc, dict.get(PDFName.of("M"))),
		};
	});
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

// Adds Highlight annotations and edits supported existing comments in a copy of
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

	const commentsById = new Map(annotations(doc).map((entry) => [entry.id, entry]));
	const seen = new Set<string>();
	for (const edit of edits) {
		if (seen.has(edit.id)) throw new Error(`comment id ${edit.id} is edited more than once`);
		seen.add(edit.id);
		if (!commentsById.has(edit.id)) throw new Error(`comment id ${edit.id} was not found`);
	}

	const date = pdfDate(new Date());
	for (const edit of edits) {
		const target = commentsById.get(edit.id);
		if (!target) throw new Error(`comment id ${edit.id} was not found`);
		target.dict.set(CONTENTS, PDFHexString.fromText(edit.comment));
		target.dict.delete(RICH_CONTENTS);
		target.dict.set(PDFName.of("M"), PDFString.of(date));
		if (target.subtype === "/FreeText") target.dict.delete(PDFName.of("AP"));
		if (target.locator) target.dict.set(ANNOTATION_NAME, PDFHexString.fromText(edit.id));
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
