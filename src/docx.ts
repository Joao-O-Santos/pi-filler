import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { dirname, join, posix, resolve } from "node:path";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { type OfficePackage, packOfficePackage, unpackOfficePackage } from "./ooxml.js";

export interface DocxPatchSet {
	page?: { width?: number; height?: number; orientation?: "portrait" | "landscape" };
	margins?: {
		top?: number;
		bottom?: number;
		left?: number;
		right?: number;
		header?: number;
		footer?: number;
		gutter?: number;
	};
	lineNumbering?: {
		mode?: "off" | "continuous" | "newPage" | "newSection";
		start?: number;
		count_by?: number;
	};
	pageNumbering?: { start?: number; format?: string };
	metadata?: {
		title?: string;
		subject?: string;
		creator?: string;
		keywords?: string;
		description?: string;
		lastModifiedBy?: string;
	};
	clearCoreMetadata?: boolean;
}

export interface DocxFormatting {
	page: { width?: number; height?: number; orientation?: string };
	margins: Record<string, number>;
	sectionCount: number;
	lineNumbering: { mode: string; start?: number; count_by?: number };
	pageNumbering: { start?: number; format?: string };
	comments: number;
	trackedChanges: number;
	metadata: Record<string, string>;
}

export interface DocxPatchResult {
	formatting: DocxFormatting;
	changedParts: string[];
	preservedParts: string[];
	written: boolean;
}

type Package = OfficePackage;
type XmlObject = { children: XmlEntry[]; attributes: Record<string, string> };
type XmlEntry = Record<string, unknown>;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";

const xmlOptions = {
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	parseTagValue: false,
	parseAttributeValue: false,
	trimValues: false,
	commentPropName: "#comment",
	cdataPropName: "#cdata",
};
const parser = new XMLParser(xmlOptions);
const builder = new XMLBuilder(xmlOptions);

function wrap(entry: XmlEntry, key: string): XmlObject {
	entry[":@"] ??= {};
	const attributes = entry[":@"] as Record<string, string>;
	return { children: entry[key] as XmlEntry[], attributes };
}

function localName(name: string): string {
	return name.replace(/^@_/, "").split(":").at(-1) ?? name;
}

function parseXml(bytes: Uint8Array | undefined, part: string): XmlObject {
	if (!bytes) throw new Error(`DOCX is missing required part ${part}`);
	try {
		return {
			children: parser.parse(new TextDecoder().decode(bytes), true) as XmlEntry[],
			attributes: {},
		};
	} catch (error) {
		throw new Error(`Invalid XML in ${part}: ${error instanceof Error ? error.message : error}`);
	}
}

function serializeXml(value: XmlObject): Uint8Array {
	return new TextEncoder().encode(builder.build(value.children));
}

function unpack(bytes: Uint8Array): Package {
	return unpackOfficePackage(bytes, "DOCX");
}

function findAll(value: XmlObject, suffix: string, found: XmlObject[] = []): XmlObject[] {
	for (const entry of value.children) {
		for (const [key, child] of Object.entries(entry)) {
			if (!Array.isArray(child)) continue;
			const node = wrap(entry, key);
			if (localName(key) === suffix) found.push(node);
			findAll(node, suffix, found);
		}
	}
	return found;
}

function first(value: XmlObject, suffix: string): XmlObject | undefined {
	return findAll(value, suffix)[0];
}

function directObject(value: XmlObject | undefined, suffix: string): XmlObject | undefined {
	if (!value) return undefined;
	for (const entry of value.children) {
		const key = Object.keys(entry).find((name) => name !== ":@" && localName(name) === suffix);
		if (key && Array.isArray(entry[key])) return wrap(entry, key);
	}
	return undefined;
}

function rootObject(value: XmlObject, suffix: string): XmlObject {
	const root = directObject(value, suffix);
	if (!root) throw new Error(`DOCX XML is missing ${suffix}`);
	return root;
}

function attr(value: XmlObject | undefined, name: string): string | undefined {
	if (!value) return undefined;
	for (const [key, child] of Object.entries(value.attributes)) {
		if (key.startsWith("@_") && localName(key) === name && child !== undefined) {
			return String(child);
		}
	}
	return undefined;
}

function numberAttr(value: XmlObject | undefined, name: string): number | undefined {
	const raw = attr(value, name);
	return raw === undefined ? undefined : Number(raw);
}

function namespacePrefix(root: XmlObject, uri: string, preferred: string): string {
	for (const [key, value] of Object.entries(root.attributes)) {
		if (key.startsWith("@_xmlns:") && value === uri) return key.slice("@_xmlns:".length);
	}
	let prefix = preferred;
	for (let suffix = 1; root.attributes[`@_xmlns:${prefix}`] !== undefined; suffix++) {
		prefix = `${preferred}${suffix}`;
	}
	root.attributes[`@_xmlns:${prefix}`] = uri;
	return prefix;
}

function setAttr(value: XmlObject, name: string, setting: number | string, prefix: string): void {
	const existing = Object.keys(value.attributes).find(
		(key) => key.startsWith("@_") && localName(key) === name,
	);
	value.attributes[existing ?? `@_${prefix}:${name}`] = String(setting);
}

function removeDirect(value: XmlObject, name: string): void {
	value.children.splice(
		0,
		value.children.length,
		...value.children.filter(
			(entry) => !Object.keys(entry).some((key) => key !== ":@" && localName(key) === name),
		),
	);
}

function countTags(value: XmlObject, suffix: string): number {
	return findAll(value, suffix).length;
}

function hasValues(value: object | undefined): boolean {
	return Boolean(value && Object.keys(value).length > 0);
}

const metadataNamespaces: Record<string, string> = {
	title: DC_NS,
	subject: DC_NS,
	creator: DC_NS,
	keywords: CORE_NS,
	description: DC_NS,
	lastModifiedBy: CORE_NS,
};

function literalText(value: XmlObject): string {
	let text = "";
	for (const entry of value.children) {
		if (entry["#text"] !== undefined) text += String(entry["#text"]);
		else
			for (const [key, child] of Object.entries(entry)) {
				if (Array.isArray(child)) text += literalText(wrap(entry, key));
			}
	}
	return text;
}

function childText(value: XmlObject, name: string): string | undefined {
	const child = directObject(value, name);
	return child ? literalText(child) : undefined;
}

function setChildText(value: XmlObject, name: string, text: string): void {
	let child = directObject(value, name);
	if (!child) {
		const uri = metadataNamespaces[name] ?? DC_NS;
		const prefix = namespacePrefix(value, uri, uri === CORE_NS ? "cp" : "dc");
		child = ensureChild(value, name, prefix);
	}
	child.children.splice(0, child.children.length, { "#text": text });
}

function coreMetadata(pkg: Package): Record<string, string> {
	const bytes = pkg["docProps/core.xml"];
	if (!bytes) return {};
	const props = rootObject(parseXml(bytes, "docProps/core.xml"), "coreProperties");
	const result: Record<string, string> = {};
	for (const name of Object.keys(metadataNamespaces)) {
		const value = childText(props, name);
		if (value !== undefined) result[name] = value;
	}
	return result;
}

function countTrackedChanges(pkg: Package, document: XmlObject): number {
	let count = countTags(document, "ins") + countTags(document, "del");
	const contentPart = /^word\/(?:header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/;
	const commentPart = relatedPart(pkg, "comments", "word/comments.xml");
	for (const [part, bytes] of Object.entries(pkg)) {
		if (!contentPart.test(part) && part !== commentPart) continue;
		const content = parseXml(bytes, part);
		count += countTags(content, "ins") + countTags(content, "del");
	}
	return count;
}

function inspectPackage(pkg: Package): DocxFormatting {
	const document = parseXml(pkg["word/document.xml"], "word/document.xml");
	const sectPrs = findAll(document, "sectPr");
	const section = sectPrs[0];
	const size = directObject(section, "pgSz");
	const margin = directObject(section, "pgMar");
	const line = directObject(section, "lnNumType");
	const pageNum = directObject(section, "pgNumType");
	const commentsPart = relatedPart(pkg, "comments", "word/comments.xml");
	return {
		page: {
			width: numberAttr(size, "w"),
			height: numberAttr(size, "h"),
			orientation: size ? (attr(size, "orient") ?? "portrait") : undefined,
		},
		margins: Object.fromEntries(
			["top", "bottom", "left", "right", "header", "footer", "gutter"]
				.map((name) => [name, numberAttr(margin, name)])
				.filter((entry): entry is [string, number] => entry[1] !== undefined),
		),
		sectionCount: sectPrs.length,
		lineNumbering: {
			mode: line ? (attr(line, "restart") ?? "newPage") : "off",
			start: numberAttr(line, "start"),
			count_by: numberAttr(line, "countBy"),
		},
		pageNumbering: { start: numberAttr(pageNum, "start"), format: attr(pageNum, "fmt") },
		comments: pkg[commentsPart]
			? countTags(parseXml(pkg[commentsPart], commentsPart), "comment")
			: 0,
		trackedChanges: countTrackedChanges(pkg, document),
		metadata: coreMetadata(pkg),
	};
}

function firstSection(document: XmlObject, prefix: string): XmlObject {
	const existing = first(document, "sectPr");
	if (existing) return existing;
	const body = first(document, "body");
	if (!body) throw new Error("DOCX is missing document body");
	return ensureChild(body, "sectPr", prefix);
}

// CT_SectPr sequence: insert new properties without moving existing children.
const sectionOrder = [
	"headerReference",
	"footerReference",
	"footnotePr",
	"endnotePr",
	"type",
	"pgSz",
	"pgMar",
	"paperSrc",
	"pgBorders",
	"lnNumType",
	"pgNumType",
	"cols",
	"formProt",
	"vAlign",
	"noEndnote",
	"titlePg",
	"textDirection",
	"bidi",
	"rtlGutter",
	"docGrid",
	"printerSettings",
	"sectPrChange",
];

function ensureChild(parent: XmlObject, name: string, prefix: string): XmlObject {
	let child = directObject(parent, name);
	if (!child) {
		const entry: XmlEntry = { [`${prefix}:${name}`]: [] };
		const rank = sectionOrder.indexOf(name);
		const next =
			rank < 0
				? -1
				: parent.children.findIndex((node) =>
						Object.keys(node).some((key) => sectionOrder.indexOf(localName(key)) > rank),
					);
		parent.children.splice(next < 0 ? parent.children.length : next, 0, entry);
		child = wrap(entry, `${prefix}:${name}`);
	}
	return child;
}

function patchPage(
	section: XmlObject,
	patch: NonNullable<DocxPatchSet["page"]>,
	prefix: string,
): void {
	const size = ensureChild(section, "pgSz", prefix);
	const width = numberAttr(size, "w");
	const height = numberAttr(size, "h");
	if (
		patch.orientation !== undefined &&
		patch.width === undefined &&
		patch.height === undefined &&
		width !== undefined &&
		height !== undefined &&
		((patch.orientation === "landscape" && width < height) ||
			(patch.orientation === "portrait" && width > height))
	) {
		setAttr(size, "w", height, prefix);
		setAttr(size, "h", width, prefix);
	}
	if (patch.width !== undefined) setAttr(size, "w", patch.width, prefix);
	if (patch.height !== undefined) setAttr(size, "h", patch.height, prefix);
	if (patch.orientation !== undefined) setAttr(size, "orient", patch.orientation, prefix);
}

function applyPatch(pkg: Package, patch: DocxPatchSet): string[] {
	const changed = new Set<string>();
	const documentPatch =
		hasValues(patch.page) ||
		hasValues(patch.margins) ||
		hasValues(patch.lineNumbering) ||
		hasValues(patch.pageNumbering);
	if (documentPatch) {
		const document = parseXml(pkg["word/document.xml"], "word/document.xml");
		const documentRoot = rootObject(document, "document");
		const prefix = namespacePrefix(documentRoot, WORD_NS, "w");
		const sectPr = firstSection(document, prefix);

		if (hasValues(patch.page)) patchPage(sectPr, patch.page ?? {}, prefix);
		if (hasValues(patch.margins)) {
			const margin = ensureChild(sectPr, "pgMar", prefix);
			for (const [name, value] of Object.entries(patch.margins ?? {})) {
				setAttr(margin, name, value, prefix);
			}
		}
		if (hasValues(patch.lineNumbering)) {
			if (patch.lineNumbering?.mode === "off") removeDirect(sectPr, "lnNumType");
			else {
				const line = ensureChild(sectPr, "lnNumType", prefix);
				if (patch.lineNumbering?.mode) {
					setAttr(line, "restart", patch.lineNumbering.mode, prefix);
				}
				if (patch.lineNumbering?.start !== undefined) {
					setAttr(line, "start", patch.lineNumbering.start, prefix);
				}
				if (patch.lineNumbering?.count_by !== undefined) {
					setAttr(line, "countBy", patch.lineNumbering.count_by, prefix);
				}
			}
		}
		if (hasValues(patch.pageNumbering)) {
			const page = ensureChild(sectPr, "pgNumType", prefix);
			if (patch.pageNumbering?.start !== undefined) {
				setAttr(page, "start", patch.pageNumbering.start, prefix);
			}
			if (patch.pageNumbering?.format !== undefined) {
				setAttr(page, "fmt", patch.pageNumbering.format, prefix);
			}
		}
		pkg["word/document.xml"] = serializeXml(document);
		changed.add("word/document.xml");
	}

	if (hasValues(patch.metadata) || patch.clearCoreMetadata) {
		const core = parseXml(pkg["docProps/core.xml"], "docProps/core.xml");
		const props = rootObject(core, "coreProperties");
		if (patch.clearCoreMetadata) {
			for (const name of Object.keys(metadataNamespaces)) {
				if (childText(props, name) !== undefined) setChildText(props, name, "");
			}
		}
		for (const [name, value] of Object.entries(patch.metadata ?? {})) {
			setChildText(props, name, value);
		}
		pkg["docProps/core.xml"] = serializeXml(core);
		changed.add("docProps/core.xml");
	}
	return [...changed];
}

function validatePatch(patch: DocxPatchSet): void {
	const requested =
		hasValues(patch.page) ||
		hasValues(patch.margins) ||
		hasValues(patch.lineNumbering) ||
		hasValues(patch.pageNumbering) ||
		hasValues(patch.metadata) ||
		patch.clearCoreMetadata === true;
	if (!requested) throw new Error("DOCX patch requires at least one change");
	if (
		patch.lineNumbering?.mode === "off" &&
		(patch.lineNumbering.start !== undefined || patch.lineNumbering.count_by !== undefined)
	) {
		throw new Error("Line numbering mode=off cannot include start or count_by");
	}
}

function validateRequested(formatting: DocxFormatting, patch: DocxPatchSet): void {
	if (patch.page) {
		for (const [key, value] of Object.entries(patch.page)) {
			if (formatting.page[key as keyof typeof formatting.page] !== value) {
				throw new Error(`DOCX page patch validation failed for ${key}`);
			}
		}
	}
	if (patch.margins) {
		for (const [key, value] of Object.entries(patch.margins)) {
			if (formatting.margins[key] !== value) {
				throw new Error(`DOCX margin patch validation failed for ${key}`);
			}
		}
	}
	if (patch.lineNumbering) {
		for (const [key, value] of Object.entries(patch.lineNumbering)) {
			if (formatting.lineNumbering[key as keyof typeof formatting.lineNumbering] !== value) {
				throw new Error(`DOCX line numbering patch validation failed for ${key}`);
			}
		}
	}
	if (patch.pageNumbering) {
		for (const [key, value] of Object.entries(patch.pageNumbering)) {
			if (formatting.pageNumbering[key as keyof typeof formatting.pageNumbering] !== value) {
				throw new Error(`DOCX page numbering patch validation failed for ${key}`);
			}
		}
	}
	if (patch.metadata) {
		for (const [key, value] of Object.entries(patch.metadata)) {
			if (formatting.metadata[key] !== value) {
				throw new Error(`DOCX metadata patch validation failed for ${key}`);
			}
		}
	}
	if (patch.clearCoreMetadata) {
		for (const name of Object.keys(metadataNamespaces)) {
			if (patch.metadata?.[name as keyof NonNullable<DocxPatchSet["metadata"]>] !== undefined) {
				continue;
			}
			if (formatting.metadata[name]) {
				throw new Error(`DOCX core metadata clearing validation failed for ${name}`);
			}
		}
	}
}

function relationshipBase(part: string): string {
	if (part.startsWith("_rels/")) return "";
	const marker = "/_rels/";
	const index = part.indexOf(marker);
	return index === -1 ? posix.dirname(part) : part.slice(0, index);
}

function validateRelationships(pkg: Package): void {
	for (const part of Object.keys(pkg).filter((name) => name.endsWith(".rels"))) {
		const root = parseXml(pkg[part], part);
		for (const relationship of findAll(root, "Relationship")) {
			const target = attr(relationship, "Target");
			if (!target || attr(relationship, "TargetMode") === "External") continue;
			const normalized = target.startsWith("/")
				? posix.normalize(target.slice(1))
				: posix.normalize(posix.join(relationshipBase(part), target));
			if (!pkg[normalized]) {
				throw new Error(`DOCX relationship target is missing: ${normalized}`);
			}
		}
	}
}

function validatePackage(pkg: Package): void {
	for (const part of ["[Content_Types].xml", "word/document.xml"]) {
		parseXml(pkg[part], part);
	}
	for (const [part, bytes] of Object.entries(pkg)) {
		if (part.endsWith(".xml") || part.endsWith(".rels")) parseXml(bytes, part);
	}
	validateRelationships(pkg);
}

async function load(path: string): Promise<Package> {
	return unpack(new Uint8Array(await readFile(resolve(path))));
}

export interface DocxReviewComment {
	id: string;
	author?: string;
	date?: string;
	text: string;
	textTruncated: boolean;
	anchor: { kind: "range" | "point" | "unmatched" | "not-found"; text: string; truncated: boolean };
	resolved?: boolean;
	replyTo?: string;
}

function reviewText(node: XmlObject): string {
	let text = "";
	for (const entry of node.children) {
		for (const [name, value] of Object.entries(entry)) {
			if (!Array.isArray(value)) continue;
			const child = wrap(entry, name);
			const tag = localName(name);
			if (tag === "t" || tag === "delText") {
				text += literalText(child);
			} else if (tag === "tab") text += "\t";
			else if (tag === "br" || tag === "cr") text += "\n";
			else text += reviewText(child);
			if (tag === "p") text += "\n";
		}
	}
	return text;
}

function relatedPart(pkg: Package, suffix: string, fallback: string): string {
	const bytes = pkg["word/_rels/document.xml.rels"];
	if (bytes) {
		for (const rel of findAll(parseXml(bytes, "word/_rels/document.xml.rels"), "Relationship")) {
			if (attr(rel, "Type")?.endsWith(`/${suffix}`) && attr(rel, "TargetMode") !== "External") {
				const target = attr(rel, "Target") ?? "";
				return posix.normalize(
					target.startsWith("/") ? target.slice(1) : posix.join("word", target),
				);
			}
		}
	}
	return fallback;
}

/** Read-only, bounded inspection; anchors include all revision text in the main document. */
export async function inspectDocxReview(path: string, offset = 0, limit = 20) {
	if (!Number.isInteger(offset) || offset < 0)
		throw new Error("review_offset must be a nonnegative integer");
	if (!Number.isInteger(limit) || limit < 1 || limit > 50)
		throw new Error("review_limit must be between 1 and 50");
	const pkg = await load(path);
	validatePackage(pkg);
	const part = relatedPart(pkg, "comments", "word/comments.xml");
	const comments = pkg[part] ? findAll(parseXml(pkg[part], part), "comment") : [];
	const selected = comments.slice(offset, offset + limit);
	const warnings = new Set<string>();
	const anchors = new Map<string, DocxReviewComment["anchor"]>();
	const active = new Set<string>();
	const ids = new Set(selected.map((comment) => attr(comment, "id") ?? ""));
	const clip = (value: string, size: number) => [...value].slice(0, size).join("");
	function append(text: string) {
		for (const id of active) {
			const anchor = anchors.get(id);
			if (!anchor) continue;
			const combined = anchor.text + text;
			anchor.text = clip(combined, 200);
			anchor.truncated ||= anchor.text !== combined;
		}
	}
	function walk(node: XmlObject, paragraph = "") {
		for (const entry of node.children) {
			for (const [name, value] of Object.entries(entry)) {
				if (!Array.isArray(value)) continue;
				const child = wrap(entry, name);
				const tag = localName(name);
				const id = attr(child, "id") ?? "";
				if (tag === "commentRangeStart" && ids.has(id)) {
					if (anchors.has(id))
						warnings.add(
							"Multiple or overlapping ranges for the same comment ID are not fully represented.",
						);
					anchors.set(id, { kind: "unmatched", text: "", truncated: false });
					active.add(id);
				} else if (tag === "commentRangeEnd" && ids.has(id)) {
					const anchor = anchors.get(id);
					if (anchor && active.has(id)) {
						anchor.kind = "range";
						active.delete(id);
					} else warnings.add("An unmatched comment range end was found.");
				} else if (tag === "commentReference" && ids.has(id) && !anchors.has(id)) {
					anchors.set(id, {
						kind: "point",
						text: clip(paragraph, 200),
						truncated: [...paragraph].length > 200,
					});
				} else if (tag === "t" || tag === "delText") {
					append(literalText(child));
				} else if (tag === "tab") append("\t");
				else if (tag === "br" || tag === "cr") append("\n");
				else walk(child, tag === "p" ? reviewText(child) : paragraph);
				if (tag === "p") append("\n");
			}
		}
	}
	walk(parseXml(pkg["word/document.xml"], "word/document.xml"));
	if (active.size) warnings.add("Some comment ranges have no closing marker.");
	const extendedPart = relatedPart(pkg, "commentsExtended", "word/commentsExtended.xml");
	const extended = pkg[extendedPart]
		? findAll(parseXml(pkg[extendedPart], extendedPart), "commentEx")
		: [];
	const paraIds = new Map(
		comments.map((comment) => [attr(findAll(comment, "p").at(-1), "paraId"), attr(comment, "id")]),
	);
	const result: DocxReviewComment[] = [];
	let consumed = 0;
	for (const comment of selected) {
		const id = attr(comment, "id") ?? "";
		if (Buffer.byteLength(id) > 512)
			throw new Error("DOCX comment identifier exceeds inspection limit");
		const fullText = reviewText(comment).replace(/\n$/, "");
		const text = clip(fullText, 1000);
		const paraId = attr(findAll(comment, "p").at(-1), "paraId");
		const ex =
			paraId === undefined ? undefined : extended.find((item) => attr(item, "paraId") === paraId);
		const done = attr(ex, "done");
		const knownDone = done !== undefined && ["1", "0", "true", "false", "on", "off"].includes(done);
		if (done !== undefined && !knownDone)
			warnings.add("An unsupported comment resolution value was found.");
		const parent = attr(ex, "paraIdParent");
		const record: DocxReviewComment = {
			id,
			author: attr(comment, "author")?.slice(0, 128),
			date: attr(comment, "date")?.slice(0, 64),
			text,
			textTruncated: text !== fullText,
			anchor: anchors.get(id) ?? { kind: "not-found", text: "", truncated: false },
			resolved: knownDone ? ["1", "true", "on"].includes(done ?? "") : undefined,
			replyTo: parent === undefined ? undefined : paraIds.get(parent),
		};
		if (parent !== undefined && record.replyTo === undefined)
			warnings.add("A comment reply parent could not be matched.");
		if (result.length && Buffer.byteLength(JSON.stringify([...result, record])) > 24000) break;
		result.push(record);
		consumed++;
	}
	if (Object.keys(pkg).some((name) => /comments(?:Ids|Extensible)|people\.xml/.test(name)))
		warnings.add(
			"Additional modern comment metadata is present; not all thread features are inspected.",
		);
	return {
		totalComments: comments.length,
		offset,
		returned: result.length,
		nextOffset: offset + consumed < comments.length ? offset + consumed : null,
		comments: result,
		warnings: [...warnings],
		anchorScope: "main document, including insertion and deletion text",
	};
}

export async function inspectDocx(path: string): Promise<DocxFormatting> {
	const pkg = await load(path);
	validatePackage(pkg);
	return inspectPackage(pkg);
}

export async function patchDocx(
	inputPath: string,
	outputPath: string,
	patch: DocxPatchSet,
	dryRun = false,
): Promise<DocxPatchResult> {
	const input = resolve(inputPath);
	const output = resolve(outputPath);
	if (input === output) throw new Error("DOCX patch output must differ from input");
	validatePatch(patch);
	const pkg = await load(input);
	validatePackage(pkg);
	const changedParts = applyPatch(pkg, patch);
	validatePackage(pkg);
	const formatting = inspectPackage(pkg);
	validateRequested(formatting, patch);
	const allParts = Object.keys(pkg).sort();
	const result = {
		formatting,
		changedParts: changedParts.sort(),
		preservedParts: allParts.filter((part) => !changedParts.includes(part)),
		written: false,
	};
	if (dryRun) return result;
	const temporary = join(dirname(output), `.${output.split(/[\\/]/).pop()}.${randomUUID()}.tmp`);
	await mkdir(dirname(output), { recursive: true });
	try {
		await writeFile(temporary, packOfficePackage(pkg));
		await inspectDocx(temporary);
		await rename(temporary, output);
		result.written = true;
		return result;
	} finally {
		await rm(temporary, { force: true });
	}
}
