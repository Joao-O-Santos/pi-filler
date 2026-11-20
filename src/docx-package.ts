import { createHash, randomUUID } from "node:crypto";
import { mkdir, readFile, realpath, rename, rm, stat, writeFile } from "node:fs/promises";
import { dirname, join, posix, resolve } from "node:path";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { type OfficePackage, packOfficePackage, unpackOfficePackage } from "./ooxml.js";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
export const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
export type Entry = Record<string, unknown>;
export interface Element {
	entry: Entry;
	name: string;
	children: Entry[];
	attrs: Record<string, string>;
	ns: Record<string, string>;
}
const options = {
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	preserveOrder: true,
	parseTagValue: false,
	parseAttributeValue: false,
	trimValues: false,
	commentPropName: "#comment",
	cdataPropName: "#cdata",
	htmlEntities: true, // Enables numeric references; parse() separately rejects non-XML names.
};
const parser = new XMLParser(options);
const builderOptions = {
	...options,
	entities: [
		{ regex: /&/g, val: "&amp;" },
		{ regex: /</g, val: "&lt;" },
		{ regex: />/g, val: "&gt;" },
		{ regex: /"/g, val: "&quot;" },
		{ regex: /'/g, val: "&apos;" },
		{ regex: /\r/g, val: "&#13;" },
	],
};
const builder = new XMLBuilder(builderOptions);

function mapAttributeWhitespace(
	source: string,
	replacement: (character: string) => string,
): string {
	return source.replace(
		/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?![!?])(?:[^<>'"]|"[^"]*"|'[^']*')+>/g,
		(tag) =>
			tag.startsWith("<!") || tag.startsWith("<?")
				? tag
				: tag.replace(/"[^"]*"|'[^']*'/g, (value) => value.replace(/[\t\n\r]/g, replacement)),
	);
}

export function elements(
	entries: Entry[],
	inherited: Record<string, string> = { xml: "http://www.w3.org/XML/1998/namespace" },
): Element[] {
	return entries.flatMap((entry) => {
		const name = Object.keys(entry).find(
			(key) => !key.startsWith("#") && key !== ":@" && !key.startsWith("?"),
		);
		if (!name || !Array.isArray(entry[name])) return [];
		const attrs = (entry[":@"] ?? {}) as Record<string, string>;
		const ns = { ...inherited };
		for (const [key, value] of Object.entries(attrs)) {
			if (key === "@_xmlns") ns[""] = value;
			else if (key.startsWith("@_xmlns:")) ns[key.slice(8)] = value;
		}
		return [{ entry, name, children: entry[name] as Entry[], attrs, ns }];
	});
}
export function uri(node: Element): string | undefined {
	return node.ns[node.name.includes(":") ? node.name.split(":")[0] : ""];
}
export function is(node: Element, namespace: string, name: string): boolean {
	return uri(node) === namespace && node.name.split(":").at(-1) === name;
}
export function children(node: Element, namespace: string, name: string): Element[] {
	return elements(node.children, node.ns).filter((child) => is(child, namespace, name));
}
export function all(node: Element): Element[] {
	return elements(node.children, node.ns).flatMap((child) => [child, ...all(child)]);
}
export function get(
	node: Element | undefined,
	namespace: string,
	name: string,
): string | undefined {
	if (!node) return undefined;
	for (const [key, value] of Object.entries(node.attrs)) {
		const qname = key.slice(2);
		const parts = qname.split(":");
		if (parts.at(-1) === name && (parts.length === 1 ? "" : node.ns[parts[0]]) === namespace)
			return value;
	}
	return undefined;
}
export function set(node: Element, namespace: string, name: string, value: string | number): void {
	const existing = Object.keys(node.attrs).find((key) => {
		const parts = key.slice(2).split(":");
		return parts.at(-1) === name && (parts.length === 1 ? "" : node.ns[parts[0]]) === namespace;
	});
	let prefix = Object.keys(node.ns).find((key) => key && node.ns[key] === namespace);
	if (namespace && !prefix) {
		prefix = "pf";
		while (node.ns[prefix]) prefix += "x";
		node.attrs[`@_xmlns:${prefix}`] = namespace;
		node.ns[prefix] = namespace;
	}
	node.attrs[existing ?? `@_${namespace ? `${prefix}:` : ""}${name}`] = String(value);
	node.entry[":@"] = node.attrs;
}
export function unset(node: Element, namespace: string, name: string): void {
	for (const key of Object.keys(node.attrs)) {
		const parts = key.slice(2).split(":");
		if (parts.at(-1) === name && (parts.length === 1 ? "" : node.ns[parts[0]]) === namespace)
			delete node.attrs[key];
	}
}
export function make(parent: Element, namespace: string, name: string): Element {
	let prefix = Object.keys(parent.ns).find((key) => parent.ns[key] === namespace);
	const attrs: Record<string, string> = {};
	if (prefix === undefined) {
		prefix = "pf";
		while (parent.ns[prefix]) prefix += "x";
		attrs[`@_xmlns:${prefix}`] = namespace;
	}
	const entry: Entry = { [prefix ? `${prefix}:${name}` : name]: [], ":@": attrs };
	return elements([entry], parent.ns)[0];
}
export function remove(parent: Element, namespace: string, name: string): void {
	const entries = new Set(children(parent, namespace, name).map((child) => child.entry));
	parent.children.splice(
		0,
		parent.children.length,
		...parent.children.filter((entry) => !entries.has(entry)),
	);
}
export function text(node: Element): string {
	return [node, ...all(node)]
		.filter((child) => is(child, W, "t"))
		.map((child) => child.children.map((entry) => entry["#text"] ?? "").join(""))
		.join("");
}
export function parse(bytes: Uint8Array | undefined, part: string): Element {
	if (!bytes) throw new Error(`DOCX is missing ${part}`);
	const source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
	if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error(`Unsupported XML declaration in ${part}`);
	assertXmlText(source);
	const entitySource = source.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, "");
	for (const match of entitySource.matchAll(/&#(x[0-9a-f]+|[0-9]+);/gi)) {
		const code = match[1].toLowerCase().startsWith("x")
			? Number.parseInt(match[1].slice(1), 16)
			: Number(match[1]);
		if (!Number.isSafeInteger(code) || code > 0x10ffff)
			throw new Error("Invalid XML character reference");
		assertXmlText(String.fromCodePoint(code));
	}
	for (const match of entitySource.matchAll(/&([a-z_][\w.:-]*);/gi))
		if (!["amp", "lt", "gt", "quot", "apos"].includes(match[1]))
			throw new Error("Unknown XML entity reference");
	// fast-xml-parser does not normalize literal attribute whitespace per XML 1.0.
	// Do so before entity decoding, leaving comments/CDATA and character references intact.
	const normalized = mapAttributeWhitespace(source.replace(/\r\n?/g, "\n"), () => " ");
	const entries = parser.parse(normalized, true) as Entry[];
	function check(values: unknown): void {
		if (typeof values === "string") assertXmlText(values);
		else if (Array.isArray(values)) values.forEach(check);
		else if (values && typeof values === "object") {
			for (const value of Object.values(values)) {
				check(value);
			}
		}
	}
	check(entries);
	const roots = elements(entries);
	if (roots.length !== 1) throw new Error(`Expected one XML root in ${part}`);
	return roots[0];
}
export function xml(root: Element): Uint8Array {
	// Keep indentation as literal whitespace for Office readers; only attributes
	// need LF/TAB escaping to survive XML attribute normalization on reopening.
	const source = mapAttributeWhitespace(
		builder.build([root.entry]),
		(character) => `&#${character.charCodeAt(0)};`,
	);
	return new TextEncoder().encode(source);
}
export function assertXmlText(value: string): void {
	if (
		[...value].some((char) => {
			const code = char.codePointAt(0) ?? 0;
			return (
				(code < 32 && ![9, 10, 13].includes(code)) ||
				(code >= 0xd800 && code <= 0xdfff) ||
				code === 0xfffe ||
				code === 0xffff
			);
		})
	)
		throw new Error("Text contains invalid XML characters");
}
export function hash(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex");
}
export function targetPart(base: string, target: string): string {
	if (/[\\?#%]/.test(target) || /^[a-z]+:/i.test(target))
		throw new Error("Unsupported relationship target URI");
	const part = posix.normalize(target.startsWith("/") ? target.slice(1) : posix.join(base, target));
	if (part.startsWith("../") || part === "..") throw new Error("Relationship escapes package");
	return part;
}
export function validate(pkg: OfficePackage): void {
	if (Object.keys(pkg).some((name) => /(^|\/)_(?:xmlsignatures)|vbaProject/i.test(name)))
		throw new Error("Signed or macro-enabled DOCX editing is unsupported");
	const document = parse(pkg["word/document.xml"], "word/document.xml");
	if (!is(document, W, "document") || children(document, W, "body").length !== 1)
		throw new Error("Unsupported DOCX document root/body");
	const types = parse(pkg["[Content_Types].xml"], "[Content_Types].xml");
	if (!is(types, CT, "Types")) throw new Error("Unsupported content-types namespace");
	if (
		elements(types.children, types.ns).some((node) =>
			/macroEnabled|vbaProject|digital-signature/i.test(get(node, "", "ContentType") ?? ""),
		)
	)
		throw new Error("Signed or macro-enabled DOCX editing is unsupported");
	for (const [name, bytes] of Object.entries(pkg)) {
		if (!/\.(xml|rels)$/.test(name)) continue;
		const root = parse(bytes, name);
		if (!name.endsWith(".rels")) continue;
		if (!is(root, REL, "Relationships")) throw new Error("Unsupported relationships namespace");
		const base = name === "_rels/.rels" ? "" : posix.dirname(name).replace(/\/_rels$/, "");
		const ids = new Set<string>();
		for (const rel of children(root, REL, "Relationship")) {
			if (/digital-signature|vbaProject/i.test(get(rel, "", "Type") ?? ""))
				throw new Error("Signed or macro-enabled DOCX editing is unsupported");
			const id = get(rel, "", "Id");
			if (!id || ids.has(id)) throw new Error("Missing or duplicate relationship ID");
			ids.add(id);
			const target = get(rel, "", "Target");
			if (!target) throw new Error("Missing relationship target");
			if (get(rel, "", "TargetMode") === "External") continue;
			if (!pkg[targetPart(base, target)])
				throw new Error("Unresolved DOCX relationship dependency");
		}
	}
}
export async function load(path: string) {
	const bytes = new Uint8Array(await readFile(path));
	const pkg = unpackOfficePackage(bytes, "DOCX");
	validate(pkg);
	return { pkg, sha256: hash(bytes) };
}
export async function distinct(output: string, inputs: string[]): Promise<void> {
	for (const input of inputs) {
		if (resolve(output) === resolve(input)) throw new Error("Output must differ from every input");
		try {
			const [a, b] = await Promise.all([stat(input), stat(output)]);
			if (a.dev === b.dev && a.ino === b.ino) throw new Error("Output aliases an input");
			if ((await realpath(input)) === (await realpath(output)))
				throw new Error("Output aliases an input");
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		}
	}
}
export async function publish(pkg: OfficePackage, output: string): Promise<void> {
	validate(pkg);
	await mkdir(dirname(output), { recursive: true });
	const temporary = join(dirname(output), `.pi-filler-${randomUUID()}.docx`);
	try {
		await writeFile(temporary, packOfficePackage(pkg));
		await load(temporary);
		await rename(temporary, output);
	} finally {
		await rm(temporary, { force: true });
	}
}
