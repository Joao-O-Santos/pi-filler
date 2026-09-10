import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm } from "node:fs/promises";
import { dirname, join, posix, resolve } from "node:path";
import { truncateHead } from "@earendil-works/pi-coding-agent";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";

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
		mode?: "off" | "continuous" | "restart" | "newPage";
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
	anonymize?: boolean;
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

type Package = Record<string, Uint8Array>;
type XmlObject = Record<string, unknown>;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	format: false,
});

function parseXml(bytes: Uint8Array, part: string): XmlObject {
	try {
		return parser.parse(new TextDecoder().decode(bytes)) as XmlObject;
	} catch (error) {
		throw new Error(`Invalid XML in ${part}: ${error instanceof Error ? error.message : error}`);
	}
}

function serializeXml(value: XmlObject): Uint8Array {
	return new TextEncoder().encode(builder.build(value));
}

function unpack(bytes: Uint8Array): Package {
	try {
		return unzipSync(bytes);
	} catch (error) {
		throw new Error(`Invalid DOCX ZIP package: ${error instanceof Error ? error.message : error}`);
	}
}

function findAll(value: unknown, suffix: string, found: XmlObject[] = []): XmlObject[] {
	if (!value || typeof value !== "object") return found;
	for (const [key, child] of Object.entries(value)) {
		if (key === suffix || key.endsWith(`:${suffix}`)) {
			for (const item of Array.isArray(child) ? child : [child]) {
				if (item && typeof item === "object") found.push(item as XmlObject);
			}
		}
		findAll(child, suffix, found);
	}
	return found;
}

function first(value: unknown, suffix: string): XmlObject | undefined {
	return findAll(value, suffix)[0];
}

function attr(value: XmlObject | undefined, name: string): string | undefined {
	if (!value) return undefined;
	return value[`@_w:${name}`] as string | undefined;
}

function numberAttr(value: XmlObject | undefined, name: string): number | undefined {
	const raw = attr(value, name);
	return raw === undefined ? undefined : Number(raw);
}

function setAttr(value: XmlObject, name: string, setting: number | string): void {
	value[`@_w:${name}`] = String(setting);
}

function remove(value: XmlObject, key: string): void {
	for (const name of Object.keys(value))
		if (name === key || name.endsWith(`:${key}`)) delete value[name];
}

function countTags(value: unknown, suffix: string): number {
	return findAll(value, suffix).length;
}

function coreMetadata(pkg: Package): Record<string, string> {
	const bytes = pkg["docProps/core.xml"];
	if (!bytes) return {};
	const root = parseXml(bytes, "docProps/core.xml");
	const props = Object.entries(root).find(([key]) => key.endsWith("coreProperties"))?.[1] as
		| XmlObject
		| undefined;
	if (!props) return {};
	const result: Record<string, string> = {};
	for (const name of ["title", "subject", "creator", "keywords", "description", "lastModifiedBy"]) {
		const value = props[`dc:${name}`] ?? props[`cp:${name}`];
		if (typeof value === "string") result[name] = value;
	}
	return result;
}

function inspectPackage(pkg: Package): DocxFormatting {
	const document = parseXml(pkg["word/document.xml"], "word/document.xml");
	const sectPrs = findAll(document, "sectPr");
	const section = sectPrs[0];
	const size = first(section, "pgSz");
	const margin = first(section, "pgMar");
	const line = first(section, "lnNumType");
	const pageNum = first(section, "pgNumType");
	const lineValue = attr(line, "restart");
	return {
		page: {
			width: numberAttr(size, "w"),
			height: numberAttr(size, "h"),
			orientation: attr(size, "orient"),
		},
		margins: Object.fromEntries(
			["top", "bottom", "left", "right", "header", "footer", "gutter"]
				.map((name) => [name, numberAttr(margin, name)])
				.filter((entry): entry is [string, number] => entry[1] !== undefined),
		),
		sectionCount: sectPrs.length,
		lineNumbering: {
			mode: line ? (lineValue === "continuous" ? "continuous" : "restart") : "off",
			start: numberAttr(line, "start"),
			count_by: numberAttr(line, "countBy"),
		},
		pageNumbering: { start: numberAttr(pageNum, "start"), format: attr(pageNum, "fmt") },
		comments: pkg["word/comments.xml"]
			? countTags(parseXml(pkg["word/comments.xml"], "word/comments.xml"), "comment")
			: 0,
		trackedChanges: countTags(document, "ins") + countTags(document, "del"),
		metadata: coreMetadata(pkg),
	};
}

function ensureObject(parent: XmlObject | undefined, suffix: string): XmlObject {
	if (!parent) throw new Error(`DOCX is missing ${suffix}`);
	let value = first(parent, suffix);
	if (!value) {
		const key = Object.keys(parent).find(
			(name) => name.endsWith(":body") || name.endsWith(":document"),
		);
		if (key) {
			value = {};
			parent[key] = { ...(parent[key] as XmlObject), [`w:${suffix}`]: value };
		}
	}
	if (!value) throw new Error(`DOCX is missing ${suffix}`);
	return value;
}

function applyPatch(pkg: Package, patch: DocxPatchSet): string[] {
	const changed = new Set<string>();
	const document = parseXml(pkg["word/document.xml"], "word/document.xml");
	const sectPr = ensureObject(document, "sectPr");
	const ensureChild = (name: string): XmlObject => {
		let child = first(sectPr, name);
		if (!child) {
			child = {};
			sectPr[`w:${name}`] = child;
		}
		return child;
	};
	if (patch.page) {
		const size = ensureChild("pgSz");
		for (const [name, value] of Object.entries(patch.page)) {
			if (name === "orientation") setAttr(size, "orient", value);
			else setAttr(size, name === "width" ? "w" : "h", value as number);
		}
		changed.add("word/document.xml");
	}
	if (patch.margins) {
		const margin = ensureChild("pgMar");
		for (const [name, value] of Object.entries(patch.margins)) setAttr(margin, name, value);
		changed.add("word/document.xml");
	}
	if (patch.lineNumbering) {
		if (patch.lineNumbering.mode === "off") remove(sectPr, "lnNumType");
		else {
			const line = ensureChild("lnNumType");
			if (patch.lineNumbering.mode)
				setAttr(
					line,
					"restart",
					patch.lineNumbering.mode === "continuous" ? "continuous" : "newPage",
				);
			if (patch.lineNumbering.start !== undefined)
				setAttr(line, "start", patch.lineNumbering.start);
			if (patch.lineNumbering.count_by !== undefined)
				setAttr(line, "countBy", patch.lineNumbering.count_by);
		}
		changed.add("word/document.xml");
	}
	if (patch.pageNumbering) {
		const page = ensureChild("pgNumType");
		if (patch.pageNumbering.start !== undefined) setAttr(page, "start", patch.pageNumbering.start);
		if (patch.pageNumbering.format !== undefined) setAttr(page, "fmt", patch.pageNumbering.format);
		changed.add("word/document.xml");
	}
	if (changed.has("word/document.xml")) pkg["word/document.xml"] = serializeXml(document);
	if (patch.metadata || patch.anonymize) {
		if (!pkg["docProps/core.xml"]) throw new Error("DOCX is missing docProps/core.xml");
		const core = parseXml(pkg["docProps/core.xml"], "docProps/core.xml");
		const props = Object.entries(core).find(([key]) =>
			key.endsWith("coreProperties"),
		)?.[1] as XmlObject;
		if (!props) throw new Error("DOCX core metadata is malformed");
		const metadata = patch.metadata ?? {};
		for (const [name, value] of Object.entries(metadata)) {
			const key = name === "lastModifiedBy" ? "cp:lastModifiedBy" : `dc:${name}`;
			props[key] = value;
		}
		if (patch.anonymize) {
			for (const key of ["dc:creator", "cp:lastModifiedBy", "dc:description", "dc:title"])
				props[key] = "";
		}
		pkg["docProps/core.xml"] = serializeXml(core);
		changed.add("docProps/core.xml");
	}
	return [...changed];
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
			if (formatting.margins[key] !== value)
				throw new Error(`DOCX margin patch validation failed for ${key}`);
		}
	}
	if (patch.lineNumbering) {
		for (const [key, value] of Object.entries(patch.lineNumbering)) {
			if (
				formatting.lineNumbering[key as keyof typeof formatting.lineNumbering] !== value &&
				!(key === "mode" && value === "restart" && formatting.lineNumbering.mode === "restart")
			) {
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
			if (formatting.metadata[key] !== value)
				throw new Error(`DOCX metadata patch validation failed for ${key}`);
		}
	}
	if (patch.anonymize && Object.values(formatting.metadata).some(Boolean)) {
		throw new Error("DOCX anonymization validation failed");
	}
}

function validatePackage(pkg: Package): void {
	for (const part of ["[Content_Types].xml", "word/document.xml"]) {
		if (!pkg[part]) throw new Error(`DOCX is missing required part ${part}`);
		parseXml(pkg[part], part);
	}
	for (const [part, bytes] of Object.entries(pkg)) {
		if (part.endsWith(".xml") || part.endsWith(".rels")) parseXml(bytes, part);
	}
	const rels = pkg["word/_rels/document.xml.rels"];
	if (rels) {
		const root = parseXml(rels, "word/_rels/document.xml.rels");
		for (const relationship of findAll(root, "Relationship")) {
			const target = relationship["@_Target"] as string | undefined;
			if (!target || relationship["@_TargetMode"] === "External") continue;
			const part = posix.normalize(posix.join("word", target));
			if (!pkg[part]) throw new Error(`DOCX relationship target is missing: ${part}`);
		}
	}
}

async function load(path: string): Promise<Package> {
	return unpack(new Uint8Array(await readFile(path)));
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
	const pkg = await load(input);
	validatePackage(pkg);
	const changedParts = applyPatch(pkg, patch);
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
		await import("node:fs/promises").then(({ writeFile }) => writeFile(temporary, zipSync(pkg)));
		await rename(temporary, output);
		result.written = true;
		return result;
	} finally {
		await rm(temporary, { force: true });
	}
}

export function summarizeFormatting(formatting: DocxFormatting): string {
	return truncateHead(JSON.stringify(formatting, null, 2)).content;
}
