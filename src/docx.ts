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
type XmlObject = Record<string, unknown>;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	format: false,
});

function isObject(value: unknown): value is XmlObject {
	return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function localName(name: string): string {
	return name.replace(/^@_/, "").split(":").at(-1) ?? name;
}

function parseXml(bytes: Uint8Array | undefined, part: string): XmlObject {
	if (!bytes) throw new Error(`DOCX is missing required part ${part}`);
	try {
		return parser.parse(new TextDecoder().decode(bytes), true) as XmlObject;
	} catch (error) {
		throw new Error(`Invalid XML in ${part}: ${error instanceof Error ? error.message : error}`);
	}
}

function serializeXml(value: XmlObject): Uint8Array {
	return new TextEncoder().encode(builder.build(value));
}

function unpack(bytes: Uint8Array): Package {
	return unpackOfficePackage(bytes, "DOCX");
}

function findAll(value: unknown, suffix: string, found: XmlObject[] = []): XmlObject[] {
	if (!value || typeof value !== "object") return found;
	for (const [key, child] of Object.entries(value)) {
		if (!key.startsWith("@_") && localName(key) === suffix) {
			for (const item of Array.isArray(child) ? child : [child]) {
				if (isObject(item)) found.push(item);
			}
		}
		findAll(child, suffix, found);
	}
	return found;
}

function first(value: unknown, suffix: string): XmlObject | undefined {
	return findAll(value, suffix)[0];
}

function directObject(value: XmlObject | undefined, suffix: string): XmlObject | undefined {
	if (!value) return undefined;
	for (const [key, child] of Object.entries(value)) {
		if (key.startsWith("@_") || localName(key) !== suffix) continue;
		const candidate = Array.isArray(child) ? child[0] : child;
		return isObject(candidate) ? candidate : undefined;
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
	for (const [key, child] of Object.entries(value)) {
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
	for (const [key, value] of Object.entries(root)) {
		if (key.startsWith("@_xmlns:") && value === uri) return key.slice("@_xmlns:".length);
	}
	root[`@_xmlns:${preferred}`] = uri;
	return preferred;
}

function setAttr(value: XmlObject, name: string, setting: number | string, prefix: string): void {
	const existing = Object.keys(value).find(
		(key) => key.startsWith("@_") && localName(key) === name,
	);
	value[existing ?? `@_${prefix}:${name}`] = String(setting);
}

function removeDirect(value: XmlObject, name: string): void {
	for (const key of Object.keys(value)) {
		if (!key.startsWith("@_") && localName(key) === name) delete value[key];
	}
}

function countTags(value: unknown, suffix: string): number {
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

function childText(value: XmlObject, name: string): string | undefined {
	for (const [key, child] of Object.entries(value)) {
		if (key.startsWith("@_") || localName(key) !== name) continue;
		if (typeof child === "string" || typeof child === "number") return String(child);
		if (isObject(child)) {
			const text = child["#text"];
			if (typeof text === "string" || typeof text === "number") return String(text);
		}
	}
	return undefined;
}

function setChildText(value: XmlObject, name: string, text: string): void {
	const existing = Object.keys(value).find(
		(key) => !key.startsWith("@_") && localName(key) === name,
	);
	if (existing) {
		const child = value[existing];
		if (isObject(child)) child["#text"] = text;
		else value[existing] = text;
		return;
	}
	const uri = metadataNamespaces[name] ?? DC_NS;
	const prefix = namespacePrefix(value, uri, uri === CORE_NS ? "cp" : "dc");
	value[`${prefix}:${name}`] = text;
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
	for (const [part, bytes] of Object.entries(pkg)) {
		if (!contentPart.test(part)) continue;
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
		comments: pkg["word/comments.xml"]
			? countTags(parseXml(pkg["word/comments.xml"], "word/comments.xml"), "comment")
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
	const section: XmlObject = {};
	body[`${prefix}:sectPr`] = section;
	return section;
}

function ensureChild(parent: XmlObject, name: string, prefix: string): XmlObject {
	let child = directObject(parent, name);
	if (!child) {
		child = {};
		parent[`${prefix}:${name}`] = child;
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
