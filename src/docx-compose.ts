import { readFile } from "node:fs/promises";
import { posix, resolve } from "node:path";
import { type Static, Type } from "typebox";
import { Check } from "typebox/value";
import {
	all,
	CT,
	children,
	distinct,
	type Element,
	elements,
	get,
	hash,
	is,
	load,
	make,
	parse,
	publish,
	R,
	REL,
	set,
	targetPart,
	text,
	unset,
	validate,
	W,
	xml,
} from "./docx-package.js";
import type { OfficePackage } from "./ooxml.js";

export const docxComposeSchema = Type.Object(
	{
		body_path: Type.String({
			minLength: 1,
			description: "Donor DOCX, not Markdown; normalized relative to working directory",
		}),
		replace_from: Type.String({
			minLength: 1,
			description:
				"Unique exact direct-body paragraph text in shell; anchor and all later body elements are replaced",
		}),
		body_from: Type.Optional(
			Type.String({
				minLength: 1,
				description:
					"Unique exact direct-body paragraph text in donor; include from this anchor, default all body elements",
			}),
		),
		expected_sha256: Type.Optional(
			Type.String({
				pattern: "^[0-9a-f]{64}$",
				description: "Optional shell input freshness hash",
			}),
		),
	},
	{
		additionalProperties: false,
		description:
			"Restricted cover clone plus body insertion. Distinct output only; imports images, hyperlinks, styles and numbering; unsupported dependencies fail closed",
	},
);
export type DocxCompose = Static<typeof docxComposeSchema>;
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

function copy(node: Element): Element {
	const result = elements([structuredClone(node.entry)], node.ns)[0];
	// Carry inherited bindings with the imported subtree, including mc:Ignorable prefixes.
	for (const [prefix, namespace] of Object.entries(node.ns))
		result.attrs[prefix ? `@_xmlns:${prefix}` : "@_xmlns"] = namespace;
	result.entry[":@"] = result.attrs;
	return result;
}
function body(root: Element): Element {
	return children(root, W, "body")[0];
}
function anchor(nodes: Element[], value: string): number {
	const matches = nodes
		.map((node, index) => (is(node, W, "p") && text(node) === value ? index : -1))
		.filter((index) => index >= 0);
	if (matches.length !== 1)
		throw new Error("Composition anchor must match exactly one direct-body paragraph");
	return matches[0];
}
function relationships(pkg: OfficePackage): Element {
	const part = "word/_rels/document.xml.rels";
	return pkg[part]
		? parse(pkg[part], part)
		: parse(new TextEncoder().encode(`<Relationships xmlns="${REL}"/>`), part);
}
function relationshipMap(root: Element): Map<string, Element> {
	return new Map(
		children(root, REL, "Relationship").map((node) => [get(node, "", "Id") ?? "", node]),
	);
}
function forbid(nodes: Element[]) {
	const unsupported = [
		"ins",
		"del",
		"moveFrom",
		"moveTo",
		"sdt",
		"altChunk",
		"sectPr",
		"footnoteReference",
		"endnoteReference",
		"commentReference",
		"commentRangeStart",
		"commentRangeEnd",
		"fldChar",
		"fldSimple",
		"instrText",
		"pPrChange",
		"rPrChange",
		"customXml",
		"pict",
		"object",
	];
	for (const node of nodes.flatMap((node) => [node, ...all(node)])) {
		if (
			get(node, W, "id") !== undefined &&
			!is(node, W, "bookmarkStart") &&
			!is(node, W, "bookmarkEnd")
		)
			throw new Error("Unresolved composition dependency: unsupported range identity");
		if (is(node, W, "drawing") && !all(node).some((child) => is(child, PIC, "pic")))
			throw new Error("Unresolved composition dependency: only picture drawings are supported");
		if (unsupported.some((name) => is(node, W, name)))
			throw new Error(`Unresolved composition dependency: unsupported ${node.name}`);
		if (is(node, W, "hyperlink") && get(node, W, "anchor"))
			throw new Error("Unresolved composition dependency: internal hyperlink anchor");
		if (node.name.split(":").at(-1)?.endsWith("Change"))
			throw new Error("Unresolved composition dependency: revision properties");
	}
}
function setStyleReferences(root: Element, mapping: Map<string, string>) {
	for (const node of [root, ...all(root)])
		if (
			["pStyle", "rStyle", "tblStyle", "basedOn", "next", "link", "styleLink", "numStyleLink"].some(
				(name) => is(node, W, name),
			)
		) {
			const old = get(node, W, "val");
			if (!old || !mapping.has(old)) throw new Error("Unresolved composition dependency: style");
			set(node, W, "val", mapping.get(old) ?? "");
		}
}
function dependencyPart(rels: Element, suffix: string): string | undefined {
	const matches = children(rels, REL, "Relationship").filter(
		(rel) => get(rel, "", "Type") === `${R}/${suffix}`,
	);
	if (matches.length > 1) throw new Error("Ambiguous composition dependency relationship");
	if (!matches.length) return undefined;
	if (get(matches[0], "", "TargetMode") === "External")
		throw new Error("External composition dependency unsupported");
	return targetPart("word", get(matches[0], "", "Target") ?? "");
}
function samePart(
	a: OfficePackage,
	ap: string | undefined,
	b: OfficePackage,
	bp: string | undefined,
): boolean {
	return (
		ap !== undefined &&
		bp !== undefined &&
		a[ap] !== undefined &&
		b[bp] !== undefined &&
		hash(a[ap]) === hash(b[bp])
	);
}
function addType(root: Element, name: string, contentType: string) {
	const existing = children(root, CT, "Override").find(
		(node) => get(node, "", "PartName") === `/${name}`,
	);
	if (existing) {
		if (get(existing, "", "ContentType") !== contentType)
			throw new Error("Conflicting composition content type");
		return;
	}
	const node = make(root, CT, "Override");
	set(node, "", "PartName", `/${name}`);
	set(node, "", "ContentType", contentType);
	root.children.push(node.entry);
}
function contentType(root: Element, part: string): string {
	const override = children(root, CT, "Override").find(
		(node) => get(node, "", "PartName") === `/${part}`,
	);
	const fallback = children(root, CT, "Default").find(
		(node) =>
			get(node, "", "Extension")?.toLowerCase() === posix.extname(part).slice(1).toLowerCase(),
	);
	const value = get(override ?? fallback, "", "ContentType");
	if (!value) throw new Error("Unresolved composition dependency: content type");
	return value;
}
export async function composeDocx(
	shellPath: string,
	outputPath: string,
	spec: DocxCompose,
	dryRun = false,
) {
	if (!Check(docxComposeSchema, spec)) throw new Error("Invalid docx_compose");
	const shell = resolve(shellPath),
		donor = resolve(spec.body_path),
		output = resolve(outputPath);
	await distinct(output, [shell, donor]);
	const { pkg, sha256 } = await load(shell);
	const source = await load(donor);
	if (spec.expected_sha256 !== undefined && spec.expected_sha256 !== sha256)
		throw new Error("Shell SHA-256 mismatch");
	const document = parse(pkg["word/document.xml"], "word/document.xml"),
		sb = body(document);
	const donorDocument = parse(source.pkg["word/document.xml"], "word/document.xml"),
		db = body(donorDocument);
	const original = elements(sb.children, sb.ns),
		donorNodes = elements(db.children, db.ns);
	const sections = original.filter((node) => is(node, W, "sectPr"));
	if (sections.length !== 1 || original.at(-1) !== sections[0])
		throw new Error("Composition requires one final shell section");
	const cut = anchor(original, spec.replace_from);
	const start = spec.body_from === undefined ? 0 : anchor(donorNodes, spec.body_from);
	const donorSections = donorNodes.filter((node) => is(node, W, "sectPr"));
	if (
		donorSections.length > 1 ||
		(donorSections.length === 1 && donorNodes.at(-1) !== donorSections[0])
	)
		throw new Error("Multiple donor sections unsupported");
	const selected = donorNodes.slice(start).filter((node) => !is(node, W, "sectPr"));
	if (!selected.length || selected.some((node) => !is(node, W, "p") && !is(node, W, "tbl")))
		throw new Error("Composition requires a nonempty paragraph/table body");
	// The shell prefix must not contain an earlier section boundary.
	if (original.slice(0, cut).some((node) => all(node).some((child) => is(child, W, "sectPr"))))
		throw new Error("Multiple shell sections unsupported");
	forbid(selected);
	const retained = original.slice(0, cut).flatMap((node) => [node, ...all(node)]);
	const retainedStarts = new Set(
		retained.filter((node) => is(node, W, "bookmarkStart")).map((node) => get(node, W, "id")),
	);
	const retainedEnds = new Set(
		retained.filter((node) => is(node, W, "bookmarkEnd")).map((node) => get(node, W, "id")),
	);
	if (
		retainedStarts.size !== retainedEnds.size ||
		[...retainedStarts].some((id) => !retainedEnds.has(id)) ||
		retained.some((node) =>
			["commentRangeStart", "commentRangeEnd", "ins", "del", "moveFrom", "moveTo"].some((name) =>
				is(node, W, name),
			),
		)
	)
		throw new Error("Composition boundary crosses a bookmark/comment/revision range");
	const removedNames = new Set(
		original
			.slice(cut)
			.flatMap((node) => [node, ...all(node)])
			.filter((node) => is(node, W, "bookmarkStart"))
			.map((node) => get(node, W, "name")),
	);
	const retainedStories = [
		...retained,
		...Object.entries(pkg)
			.filter(([name]) => /^word\/(header|footer)\d+\.xml$/.test(name))
			.flatMap(([name, bytes]) => all(parse(bytes, name))),
	];
	if (
		retainedStories.some(
			(node) => is(node, W, "hyperlink") && removedNames.has(get(node, W, "anchor")),
		)
	)
		throw new Error("Retained hyperlink targets removed body bookmark");
	if (
		retainedStories.some(
			(node) =>
				is(node, W, "instrText") &&
				/\b(?:REF|PAGEREF|NOTEREF)\b/i.test(
					node.children.map((entry) => entry["#text"] ?? "").join(""),
				),
		)
	)
		throw new Error("Retained cross-reference fields are unsupported");
	const imported = selected.map(copy);
	const rels = relationships(pkg),
		sr = relationships(source.pkg),
		byId = relationshipMap(sr);
	const types = parse(pkg["[Content_Types].xml"], "[Content_Types].xml"),
		stypes = parse(source.pkg["[Content_Types].xml"], "[Content_Types].xml");
	const changed = new Set([
		"word/document.xml",
		"word/_rels/document.xml.rels",
		"[Content_Types].xml",
	]);
	const ids = new Set(relationshipMap(rels).keys());
	let serial = 1;
	function newRel(type: string, target: string, external = false): string {
		let id = `rIdFiller${serial++}`;
		while (ids.has(id)) id = `rIdFiller${serial++}`;
		ids.add(id);
		const node = make(rels, REL, "Relationship");
		for (const [key, value] of Object.entries({ Id: id, Type: type, Target: target }))
			set(node, "", key, value);
		if (external) set(node, "", "TargetMode", "External");
		rels.children.push(node.entry);
		return id;
	}
	function freshPart(stem: string, extension: string): string {
		let n = 1;
		while (pkg[`word/${stem}${n}${extension}`]) n++;
		return `word/${stem}${n}${extension}`;
	}
	const relMapping = new Map<string, string>();
	for (const node of imported.flatMap((node) => [node, ...all(node)])) {
		for (const key of Object.keys(node.attrs)) {
			const parts = key.slice(2).split(":");
			if (parts.length !== 2 || node.ns[parts[0]] !== R) continue;
			const old = node.attrs[key];
			if (!relMapping.has(old)) {
				const rel = byId.get(old),
					type = get(rel, "", "Type"),
					target = get(rel, "", "Target");
				if (!rel || !target || !type)
					throw new Error("Unresolved composition dependency: relationship ID");
				const external = get(rel, "", "TargetMode") === "External";
				if (type === `${R}/hyperlink` && external) {
					if (!/^(https?:|mailto:)/i.test(target))
						throw new Error("Unsupported external hyperlink scheme");
					relMapping.set(old, newRel(type, target, true));
				} else if (type === `${R}/image` && !external) {
					const part = targetPart("word", target);
					const mime = contentType(stypes, part);
					if (!/^image\/(png|jpeg|gif|bmp|tiff|x-emf|x-wmf|svg\+xml)$/.test(mime))
						throw new Error("Unsupported image content type");
					if (source.pkg[`${posix.dirname(part)}/_rels/${posix.basename(part)}.rels`])
						throw new Error("Unresolved composition dependency: image relationships");
					const dest = freshPart("media/filler", posix.extname(part));
					pkg[dest] = source.pkg[part];
					changed.add(dest);
					addType(types, dest, mime);
					relMapping.set(old, newRel(type, posix.relative("word", dest)));
				} else
					throw new Error(
						"Unresolved composition dependency: only internal images and external hyperlinks are supported",
					);
			}
			node.attrs[key] = relMapping.get(old) ?? "";
		}
	}
	// Import donor styles under fresh IDs. A synthetic base materializes donor docDefaults
	// without changing the shell's defaults or styles used by its cover and headers.
	const stylesPart = dependencyPart(sr, "styles");
	const shellStylesPart = dependencyPart(rels, "styles");
	const usedStyleIds = imported
		.flatMap((node) => [node, ...all(node)])
		.filter((node) => ["pStyle", "rStyle", "tblStyle"].some((name) => is(node, W, name)));
	if (!stylesPart && usedStyleIds.length)
		throw new Error("Unresolved composition dependency: styles part");
	const styleMapping = new Map<string, string>();
	let styleCount = 0;
	let stylesNeedNumbering = false;
	const importedStyleIds = new Set<string>();
	if (stylesPart) {
		const styles = parse(source.pkg[stylesPart], stylesPart);
		const destinationPart = shellStylesPart ?? freshPart("fillerStyles", ".xml");
		const dest = shellStylesPart
			? parse(pkg[shellStylesPart], shellStylesPart)
			: parse(new TextEncoder().encode(`<w:styles xmlns:w="${W}"/>`), destinationPart);
		const availableStyles = children(styles, W, "style");
		const defaultStyles = availableStyles.filter(
			(style) =>
				get(style, W, "type") === "paragraph" &&
				["1", "true", "on"].includes(get(style, W, "default") ?? ""),
		);
		if (defaultStyles.length > 1) throw new Error("Ambiguous default paragraph style");
		const required = new Set(usedStyleIds.map((node) => get(node, W, "val") ?? ""));
		if (defaultStyles[0]) required.add(get(defaultStyles[0], W, "styleId") ?? "");
		const numberingDependency = dependencyPart(sr, "numbering");
		if (numberingDependency) {
			const numbering = parse(source.pkg[numberingDependency], numberingDependency);
			for (const node of all(numbering))
				if (["pStyle", "rStyle", "styleLink", "numStyleLink"].some((name) => is(node, W, name)))
					required.add(get(node, W, "val") ?? "");
		}
		const sourceStyles: Element[] = [];
		for (const id of required) {
			const matches = availableStyles.filter((style) => get(style, W, "styleId") === id);
			if (matches.length !== 1)
				throw new Error(
					`Unresolved composition dependency: missing or ambiguous used style ID ${id.slice(0, 128)}`,
				);
			sourceStyles.push(matches[0]);
			for (const node of all(matches[0]))
				if (["basedOn", "next", "link"].some((name) => is(node, W, name)))
					required.add(get(node, W, "val") ?? "");
		}
		const taken = new Set(children(dest, W, "style").map((node) => get(node, W, "styleId")));
		let index = 1;
		const nextId = () => {
			let id = `FillerStyle${index++}`;
			while (taken.has(id)) id = `FillerStyle${index++}`;
			taken.add(id);
			return id;
		};
		const baseId = nextId();
		importedStyleIds.add(baseId);
		for (const style of sourceStyles) {
			const id = get(style, W, "styleId");
			if (!id || styleMapping.has(id)) throw new Error("Missing or duplicate donor style ID");
			styleMapping.set(id, nextId());
		}
		const defaults = children(styles, W, "docDefaults")[0];
		const styleNodes = [...sourceStyles, ...(defaults ? [defaults] : [])];
		stylesNeedNumbering = styleNodes.some((node) =>
			all(node).some((child) => is(child, W, "numId") && get(child, W, "val") !== "0"),
		);
		for (const id of styleMapping.values()) importedStyleIds.add(id);
		const themed = [...styleNodes, ...imported].some((node) =>
			[node, ...all(node)].some((child) =>
				Object.keys(child.attrs).some((key) => /Theme|theme/.test(key)),
			),
		);
		if (
			themed &&
			!samePart(pkg, dependencyPart(rels, "theme"), source.pkg, dependencyPart(sr, "theme"))
		)
			throw new Error("Unresolved composition dependency: differing or missing themes");
		const base = make(dest, W, "style");
		set(base, W, "type", "paragraph");
		set(base, W, "styleId", baseId);
		for (const [wrapper, name] of [
			["pPrDefault", "pPr"],
			["rPrDefault", "rPr"],
		]) {
			const value = defaults && children(defaults, W, wrapper)[0];
			const props = value && children(value, W, name)[0];
			if (props) base.children.push(copy(props).entry);
		}
		dest.children.push(base.entry);
		for (const style of sourceStyles) {
			const cloned = copy(style),
				old = get(style, W, "styleId") ?? "";
			set(cloned, W, "styleId", styleMapping.get(old) ?? "");
			unset(cloned, W, "default");
			setStyleReferences(cloned, styleMapping);
			if (get(cloned, W, "type") === "paragraph" && !children(cloned, W, "basedOn").length) {
				const based = make(cloned, W, "basedOn");
				set(based, W, "val", baseId);
				const before = elements(cloned.children, cloned.ns).find(
					(node) => !is(node, W, "name") && !is(node, W, "aliases"),
				);
				cloned.children.splice(
					before ? cloned.children.indexOf(before.entry) : cloned.children.length,
					0,
					based.entry,
				);
			}
			dest.children.push(cloned.entry);
		}
		for (const node of imported) setStyleReferences(node, styleMapping);
		const defaultStyle = sourceStyles.find(
			(style) =>
				get(style, W, "type") === "paragraph" &&
				["1", "true", "on"].includes(get(style, W, "default") ?? ""),
		);
		for (const p of imported
			.flatMap((node) => [node, ...all(node)])
			.filter((node) => is(node, W, "p"))) {
			let pp = children(p, W, "pPr")[0];
			if (!pp) {
				pp = make(p, W, "pPr");
				p.children.unshift(pp.entry);
			}
			if (!children(pp, W, "pStyle").length) {
				const ps = make(pp, W, "pStyle");
				set(
					ps,
					W,
					"val",
					defaultStyle
						? (styleMapping.get(get(defaultStyle, W, "styleId") ?? "") ?? baseId)
						: baseId,
				);
				pp.children.unshift(ps.entry);
			}
		}
		pkg[destinationPart] = xml(dest);
		changed.add(destinationPart);
		styleCount = sourceStyles.length + 1;
		if (!shellStylesPart) {
			newRel(`${R}/styles`, posix.relative("word", destinationPart));
			addType(types, destinationPart, contentType(stypes, stylesPart));
		}
	}
	// Numbering imports preserve abstract definitions, overrides and style links.
	const numPart = dependencyPart(sr, "numbering");
	const donorNumIds = imported
		.flatMap((node) => [node, ...all(node)])
		.filter((node) => is(node, W, "numId"));
	let numberingCount = 0;
	if (donorNumIds.some((node) => get(node, W, "val") !== "0") || stylesNeedNumbering) {
		if (!numPart) throw new Error("Unresolved composition dependency: numbering");
	}
	if (numPart) {
		const numbering = parse(source.pkg[numPart], numPart);
		if (
			all(numbering).some(
				(node) =>
					is(node, W, "numPicBullet") ||
					Object.keys(node.attrs).some(
						(key) =>
							key.slice(2).split(":").length === 2 && node.ns[key.slice(2).split(":")[0]] === R,
					),
			)
		)
			throw new Error("Unresolved composition dependency: picture numbering");
		const existing = dependencyPart(rels, "numbering"),
			destPart = existing ?? freshPart("fillerNumbering", ".xml");
		const dest = existing
			? parse(pkg[existing], existing)
			: parse(new TextEncoder().encode(`<w:numbering xmlns:w="${W}"/>`), destPart);
		const numMap = new Map<string, string>(),
			abstractMap = new Map<string, string>();
		for (const [name, attrName, mapping] of [
			["abstractNum", "abstractNumId", abstractMap],
			["num", "numId", numMap],
		] as const) {
			const taken = new Set(children(dest, W, name).map((node) => get(node, W, attrName)));
			let n = 1;
			for (const node of children(numbering, W, name)) {
				const id = get(node, W, attrName);
				if (!id || mapping.has(id)) throw new Error("Ambiguous numbering IDs");
				while (taken.has(String(n))) n++;
				mapping.set(id, String(n));
				taken.add(String(n++));
			}
		}
		const copied = elements(numbering.children, numbering.ns).map(copy);
		for (const node of copied) {
			if (!is(node, W, "abstractNum") && !is(node, W, "num"))
				throw new Error("Unsupported numbering metadata");
			const name = is(node, W, "num") ? "numId" : "abstractNumId",
				mapping = name === "numId" ? numMap : abstractMap;
			set(node, W, name, mapping.get(get(node, W, name) ?? "") ?? "");
			setStyleReferences(node, styleMapping);
			for (const ref of all(node).filter((child) => is(child, W, "abstractNumId"))) {
				const mapped = abstractMap.get(get(ref, W, "val") ?? "");
				if (!mapped) throw new Error("Unresolved numbering abstract ID");
				set(ref, W, "val", mapped);
			}
			// All abstract definitions precede concrete nums.
			const firstNum = children(dest, W, "num")[0];
			dest.children.splice(
				is(node, W, "abstractNum") && firstNum
					? dest.children.indexOf(firstNum.entry)
					: dest.children.length,
				0,
				node.entry,
			);
		}
		function remapNums(root: Element) {
			for (const node of [root, ...all(root)].filter((child) => is(child, W, "numId"))) {
				const id = get(node, W, "val") ?? "";
				if (id === "0") continue;
				const mapped = numMap.get(id);
				if (!mapped) throw new Error("Unresolved numbering ID");
				set(node, W, "val", mapped);
			}
		}
		for (const node of imported) remapNums(node);
		// Only newly imported styles are remapped; existing shell styles stay untouched.
		const destStylesPart = dependencyPart(rels, "styles");
		if (destStylesPart) {
			const styles = parse(pkg[destStylesPart], destStylesPart);
			for (const node of children(styles, W, "style"))
				if (importedStyleIds.has(get(node, W, "styleId") ?? "")) remapNums(node);
			pkg[destStylesPart] = xml(styles);
		}
		pkg[destPart] = xml(dest);
		changed.add(destPart);
		numberingCount = numMap.size;
		if (!existing) {
			newRel(`${R}/numbering`, posix.relative("word", destPart));
			addType(types, destPart, contentType(stypes, numPart));
		}
	}
	// IDs shared across stories: drawings and bookmarks receive fresh values.
	const allShell = Object.entries(pkg)
		.filter(([name]) => /^word\/.*\.xml$/.test(name))
		.flatMap(([name, bytes]) => {
			const root = parse(bytes, name);
			return [root, ...all(root)];
		});
	const drawingIds = new Set(
		allShell
			.filter((node) => is(node, WP, "docPr") || is(node, PIC, "cNvPr"))
			.map((node) => String(Number(get(node, "", "id")))),
	);
	const bookmarkIds = new Set(
		allShell
			.filter((node) => is(node, W, "bookmarkStart"))
			.map((node) => String(Number(get(node, W, "id")))),
	);
	const bookmarkNames = new Set(
		allShell.filter((node) => is(node, W, "bookmarkStart")).map((node) => get(node, W, "name")),
	);
	let drawingId = 1,
		bookmarkId = 1;
	const bookmarks = new Map<string, string>();
	for (const node of imported.flatMap((node) => [node, ...all(node)])) {
		if (is(node, WP, "docPr") || is(node, PIC, "cNvPr")) {
			while (drawingIds.has(String(drawingId))) drawingId++;
			if (drawingId > 0xffffffff) throw new Error("Drawing ID space exhausted");
			set(node, "", "id", drawingId);
			drawingIds.add(String(drawingId++));
		}
		if (is(node, W, "bookmarkStart")) {
			const old = get(node, W, "id") ?? "";
			if (bookmarks.has(old)) throw new Error("Duplicate donor bookmark ID");
			while (
				bookmarkIds.has(String(bookmarkId)) ||
				bookmarkNames.has(`FillerBookmark${bookmarkId}`)
			)
				bookmarkId++;
			if (bookmarkId > 0x7fffffff) throw new Error("Bookmark ID space exhausted");
			const id = String(bookmarkId++);
			bookmarkIds.add(id);
			bookmarkNames.add(`FillerBookmark${id}`);
			bookmarks.set(old, id);
			set(node, W, "id", id);
			set(node, W, "name", `FillerBookmark${id}`);
		}
	}
	const closedBookmarks = new Set<string>();
	const openBookmarks = new Set<string>();
	for (const node of imported.flatMap((node) => [node, ...all(node)])) {
		if (is(node, W, "bookmarkStart")) openBookmarks.add(get(node, W, "id") ?? "");
		if (is(node, W, "bookmarkEnd")) {
			const mapped = bookmarks.get(get(node, W, "id") ?? "");
			if (!mapped || !openBookmarks.has(mapped) || closedBookmarks.has(mapped))
				throw new Error("Unmatched or duplicate donor bookmark end");
			closedBookmarks.add(mapped);
			set(node, W, "id", mapped);
		}
		for (const key of Object.keys(node.attrs))
			if (/:(paraId|textId|anchorId|editId)$/.test(key)) delete node.attrs[key];
	}
	if (closedBookmarks.size !== bookmarks.size) throw new Error("Unclosed donor bookmark");
	// Remove all old suffix nodes, not just visible text; retain only the final section.
	const cutEntry = original[cut].entry;
	sb.children.splice(
		sb.children.indexOf(cutEntry),
		sb.children.length,
		...imported.map((node) => node.entry),
		sections[0].entry,
	);
	pkg["word/document.xml"] = xml(document);
	pkg["word/_rels/document.xml.rels"] = xml(rels);
	pkg["[Content_Types].xml"] = xml(types);
	if (
		hash(new Uint8Array(await readFile(shell))) !== sha256 ||
		hash(new Uint8Array(await readFile(donor))) !== source.sha256
	)
		throw new Error("Composition input changed during processing");
	validate(pkg);
	if (!dryRun) await publish(pkg, output);
	return {
		sha256,
		bodySha256: source.sha256,
		keptElements: cut,
		insertedElements: imported.length,
		importedRelationships: relMapping.size,
		importedStyles: styleCount,
		importedNumbering: numberingCount,
		changedParts: [...changed].sort(),
		unresolvedDependencies: [],
		written: !dryRun,
		warnings: [
			"Restricted OOXML composition, not visual equivalence. Shell headers, footers and final section retained. Render to check page breaks; review old ancillary package parts before distributing confidential templates.",
		],
	};
}
