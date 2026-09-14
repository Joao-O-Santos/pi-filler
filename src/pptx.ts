import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { dirname, join, posix, resolve } from "node:path";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { type OfficePackage, packOfficePackage, unpackOfficePackage } from "./ooxml.js";

export interface PptxPatchSet {
	slideSize?: { width?: number; height?: number };
	metadata?: {
		title?: string;
		subject?: string;
		creator?: string;
		keywords?: string;
		description?: string;
		lastModifiedBy?: string;
	};
	replaceText?: Array<{ find: string; replace: string; slide?: number }>;
	clearCoreMetadata?: boolean;
}

export interface PptxFormatting {
	slideSize: { width?: number; height?: number };
	slideCount: number;
	metadata: Record<string, string>;
}

export interface PptxPatchResult {
	formatting: PptxFormatting;
	replacementCount: number;
	changedParts: string[];
	preservedParts: string[];
	written: boolean;
}

type Package = OfficePackage;
type XmlObject = Record<string, unknown>;

type TextNode = { parent: XmlObject; key: string; text: string };

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const metadataNamespaces: Record<string, string> = {
	title: DC_NS,
	subject: DC_NS,
	creator: DC_NS,
	keywords: CORE_NS,
	description: DC_NS,
	lastModifiedBy: CORE_NS,
};

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	trimValues: false,
});
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
	if (!bytes) throw new Error(`PPTX is missing required part ${part}`);
	try {
		return parser.parse(new TextDecoder().decode(bytes), true) as XmlObject;
	} catch (error) {
		throw new Error(`Invalid XML in ${part}: ${error instanceof Error ? error.message : error}`);
	}
}

function serializeXml(value: XmlObject): Uint8Array {
	return new TextEncoder().encode(builder.build(value));
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

function findPrefixed(
	value: unknown,
	prefix: string,
	suffix: string,
	found: XmlObject[] = [],
): XmlObject[] {
	if (!value || typeof value !== "object") return found;
	for (const [key, child] of Object.entries(value)) {
		if (!key.startsWith("@_") && key === `${prefix}:${suffix}`) {
			for (const item of Array.isArray(child) ? child : [child]) {
				if (isObject(item)) found.push(item);
			}
		}
		findPrefixed(child, prefix, suffix, found);
	}
	return found;
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
	if (!root) throw new Error(`PPTX XML is missing ${suffix}`);
	return root;
}

function attr(value: XmlObject | undefined, name: string): string | undefined {
	if (!value) return undefined;
	for (const [key, child] of Object.entries(value)) {
		if (key.startsWith("@_") && localName(key) === name && child !== undefined)
			return String(child);
	}
	return undefined;
}

function numberAttr(value: XmlObject | undefined, name: string): number | undefined {
	const raw = attr(value, name);
	return raw === undefined ? undefined : Number(raw);
}

function existingNamespacePrefix(root: XmlObject, uri: string): string | undefined {
	for (const [key, value] of Object.entries(root)) {
		if (key === "@_xmlns" && value === uri) return "";
		if (key.startsWith("@_xmlns:") && value === uri) return key.slice("@_xmlns:".length);
	}
	return undefined;
}

function namespacePrefix(root: XmlObject, uri: string, preferred: string): string {
	return (
		existingNamespacePrefix(root, uri) ??
		(() => {
			root[`@_xmlns:${preferred}`] = uri;
			return preferred;
		})()
	);
}

function setAttr(value: XmlObject, name: string, setting: number | string, prefix?: string): void {
	const existing = Object.keys(value).find(
		(key) => key.startsWith("@_") && localName(key) === name,
	);
	value[existing ?? `@_${prefix ? `${prefix}:` : ""}${name}`] = String(setting);
}

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
	value[`${prefix ? `${prefix}:` : ""}${name}`] = text;
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

function relationshipBase(part: string): string {
	if (part.startsWith("_rels/")) return "";
	const marker = "/_rels/";
	const index = part.indexOf(marker);
	return index === -1 ? posix.dirname(part) : part.slice(0, index);
}

function relationshipTarget(part: string, target: string): string {
	return target.startsWith("/")
		? posix.normalize(target.slice(1))
		: posix.normalize(posix.join(relationshipBase(part), target));
}

function validateRelationships(pkg: Package): void {
	for (const part of Object.keys(pkg).filter((name) => name.endsWith(".rels"))) {
		const root = parseXml(pkg[part], part);
		for (const relationship of findAll(root, "Relationship")) {
			const target = attr(relationship, "Target");
			if (!target || attr(relationship, "TargetMode") === "External") continue;
			const normalized = relationshipTarget(part, target);
			if (!pkg[normalized]) throw new Error(`PPTX relationship target is missing: ${normalized}`);
		}
	}
}

function validatePackage(pkg: Package): void {
	for (const part of [
		"[Content_Types].xml",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
	]) {
		parseXml(pkg[part], part);
	}
	for (const [part, bytes] of Object.entries(pkg)) {
		if (part.endsWith(".xml") || part.endsWith(".rels")) parseXml(bytes, part);
	}
	validateRelationships(pkg);
}

function slideParts(pkg: Package, presentation: XmlObject): string[] {
	const relsPart = "ppt/_rels/presentation.xml.rels";
	const relationships = findAll(parseXml(pkg[relsPart], relsPart), "Relationship");
	const slides = new Map<string, string>();
	for (const relationship of relationships) {
		const type = attr(relationship, "Type");
		const id = attr(relationship, "Id");
		const target = attr(relationship, "Target");
		if (type?.endsWith("/slide") && id && target)
			slides.set(id, relationshipTarget(relsPart, target));
	}
	const ordered = findAll(presentation, "sldId")
		.map((slide) => attr(slide, "id"))
		.filter((id): id is string => id !== undefined)
		.map((id) => slides.get(id))
		.filter((part): part is string => part !== undefined);
	return ordered.length > 0 ? ordered : [...slides.values()];
}

function inspectPackage(pkg: Package): PptxFormatting {
	const presentation = parseXml(pkg["ppt/presentation.xml"], "ppt/presentation.xml");
	const root = rootObject(presentation, "presentation");
	const size = directObject(root, "sldSz");
	return {
		slideSize: { width: numberAttr(size, "cx"), height: numberAttr(size, "cy") },
		slideCount: slideParts(pkg, presentation).length,
		metadata: coreMetadata(pkg),
	};
}

function textNodes(value: unknown, found: TextNode[] = []): TextNode[] {
	if (!value || typeof value !== "object") return found;
	for (const [key, child] of Object.entries(value)) {
		if (key.startsWith("@_")) continue;
		if (localName(key) === "t" && (typeof child === "string" || typeof child === "number")) {
			found.push({ parent: value as XmlObject, key, text: String(child) });
		} else {
			textNodes(child, found);
		}
	}
	return found;
}

function replaceInParagraph(paragraph: XmlObject, find: string, replacement: string): number {
	const nodes = textNodes(paragraph);
	let full = nodes.map((node) => node.text).join("");
	let count = 0;
	let start = full.indexOf(find);
	while (start !== -1) {
		const end = start + find.length;
		let cursor = 0;
		let startNode: TextNode | undefined;
		let endNode: TextNode | undefined;
		let startOffset = 0;
		let endOffset = 0;
		for (const node of nodes) {
			const nodeEnd = cursor + node.text.length;
			if (!startNode && start >= cursor && start <= nodeEnd) {
				startNode = node;
				startOffset = start - cursor;
			}
			if (end >= cursor && end <= nodeEnd) {
				endNode = node;
				endOffset = end - cursor;
				break;
			}
			cursor = nodeEnd;
		}
		if (!startNode || !endNode) break;
		if (startNode === endNode) {
			startNode.text =
				startNode.text.slice(0, startOffset) + replacement + startNode.text.slice(endOffset);
			startNode.parent[startNode.key] = startNode.text;
		} else {
			startNode.text = startNode.text.slice(0, startOffset) + replacement;
			startNode.parent[startNode.key] = startNode.text;
			let clearing = false;
			for (const node of nodes) {
				if (node === startNode) {
					clearing = true;
					continue;
				}
				if (node === endNode) {
					node.text = node.text.slice(endOffset);
					node.parent[node.key] = node.text;
					break;
				}
				if (clearing) {
					node.text = "";
					node.parent[node.key] = "";
				}
			}
		}
		full = nodes.map((node) => node.text).join("");
		count += 1;
		start = full.indexOf(find, start + replacement.length);
	}
	return count;
}

function replaceInSlide(document: XmlObject, find: string, replacement: string): number {
	const root = rootObject(document, "sld");
	const drawingPrefix = existingNamespacePrefix(root, DRAWING_NS) ?? "a";
	let count = 0;
	for (const paragraph of findPrefixed(document, drawingPrefix, "p")) {
		count += replaceInParagraph(paragraph, find, replacement);
	}
	return count;
}

function applyPatch(
	pkg: Package,
	patch: PptxPatchSet,
): { changedParts: string[]; replacementCount: number } {
	const changed = new Set<string>();
	let replacementCount = 0;
	const presentation = parseXml(pkg["ppt/presentation.xml"], "ppt/presentation.xml");
	const root = rootObject(presentation, "presentation");
	if (patch.slideSize) {
		const size =
			directObject(root, "sldSz") ??
			(() => {
				const value: XmlObject = {};
				const prefix = existingNamespacePrefix(root, PRESENTATION_NS) ?? "p";
				root[`${prefix ? `${prefix}:` : ""}sldSz`] = value;
				return value;
			})();
		if (patch.slideSize.width !== undefined) setAttr(size, "cx", patch.slideSize.width);
		if (patch.slideSize.height !== undefined) setAttr(size, "cy", patch.slideSize.height);
		pkg["ppt/presentation.xml"] = serializeXml(presentation);
		changed.add("ppt/presentation.xml");
	}

	if (patch.metadata || patch.clearCoreMetadata) {
		const core = parseXml(pkg["docProps/core.xml"], "docProps/core.xml");
		const props = rootObject(core, "coreProperties");
		if (patch.clearCoreMetadata) {
			for (const name of Object.keys(metadataNamespaces)) {
				if (childText(props, name) !== undefined) setChildText(props, name, "");
			}
		}
		for (const [name, value] of Object.entries(patch.metadata ?? {}))
			setChildText(props, name, value);
		pkg["docProps/core.xml"] = serializeXml(core);
		changed.add("docProps/core.xml");
	}

	if (patch.replaceText) {
		const parts = slideParts(pkg, presentation);
		const documents = new Map<string, XmlObject>();
		for (const request of patch.replaceText) {
			const targets = request.slide === undefined ? parts : [parts[request.slide - 1]];
			if (targets.some((part) => part === undefined))
				throw new Error("PPTX text replacement slide is out of range");
			let matches = 0;
			for (const part of targets) {
				let document = documents.get(part);
				if (!document) {
					document = parseXml(pkg[part], part);
					documents.set(part, document);
				}
				matches += replaceInSlide(document, request.find, request.replace);
			}
			if (matches === 0) throw new Error("PPTX text replacement found no matches");
			replacementCount += matches;
		}
		for (const [part, document] of documents) {
			pkg[part] = serializeXml(document);
			changed.add(part);
		}
	}
	return { changedParts: [...changed].sort(), replacementCount };
}

function hasValues(value: object | undefined): boolean {
	return Boolean(value && Object.keys(value).length > 0);
}

function validatePatch(patch: PptxPatchSet): void {
	if (
		!hasValues(patch.slideSize) &&
		!hasValues(patch.metadata) &&
		!patch.replaceText?.length &&
		patch.clearCoreMetadata !== true
	) {
		throw new Error("PPTX patch requires at least one change");
	}
	if (patch.replaceText?.some((request) => !request.find)) {
		throw new Error("PPTX text replacements require non-empty find values");
	}
}

function validateRequested(formatting: PptxFormatting, patch: PptxPatchSet): void {
	if (patch.slideSize) {
		for (const [key, value] of Object.entries(patch.slideSize)) {
			if (formatting.slideSize[key as keyof typeof formatting.slideSize] !== value) {
				throw new Error(`PPTX slide size patch validation failed for ${key}`);
			}
		}
	}
	if (patch.metadata) {
		for (const [key, value] of Object.entries(patch.metadata)) {
			if (formatting.metadata[key] !== value)
				throw new Error(`PPTX metadata patch validation failed for ${key}`);
		}
	}
	if (patch.clearCoreMetadata) {
		for (const name of Object.keys(metadataNamespaces)) {
			if (patch.metadata?.[name as keyof NonNullable<PptxPatchSet["metadata"]>] !== undefined)
				continue;
			if (formatting.metadata[name])
				throw new Error(`PPTX core metadata clearing validation failed for ${name}`);
		}
	}
}

async function load(path: string): Promise<Package> {
	return unpackOfficePackage(new Uint8Array(await readFile(resolve(path))), "PPTX");
}

export async function inspectPptx(path: string): Promise<PptxFormatting> {
	const pkg = await load(path);
	validatePackage(pkg);
	return inspectPackage(pkg);
}

export async function patchPptx(
	inputPath: string,
	outputPath: string,
	patch: PptxPatchSet,
	dryRun = false,
): Promise<PptxPatchResult> {
	const input = resolve(inputPath);
	const output = resolve(outputPath);
	if (input === output) throw new Error("PPTX patch output must differ from input");
	validatePatch(patch);
	const pkg = await load(input);
	validatePackage(pkg);
	const applied = applyPatch(pkg, patch);
	validatePackage(pkg);
	const formatting = inspectPackage(pkg);
	validateRequested(formatting, patch);
	const allParts = Object.keys(pkg).sort();
	const result: PptxPatchResult = {
		formatting,
		replacementCount: applied.replacementCount,
		changedParts: applied.changedParts,
		preservedParts: allParts.filter((part) => !applied.changedParts.includes(part)),
		written: false,
	};
	if (dryRun) return result;
	const temporary = join(dirname(output), `.${output.split(/[\\/]/).pop()}.${randomUUID()}.tmp`);
	await mkdir(dirname(output), { recursive: true });
	try {
		await writeFile(temporary, packOfficePackage(pkg));
		await inspectPptx(temporary);
		await rename(temporary, output);
		result.written = true;
		return result;
	} finally {
		await rm(temporary, { force: true });
	}
}
