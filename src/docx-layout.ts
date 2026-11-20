import { resolve } from "node:path";
import { type Static, Type } from "typebox";
import { Check } from "typebox/value";
import {
	all,
	assertXmlText,
	children,
	distinct,
	type Element,
	elements,
	get,
	is,
	load,
	make,
	parse,
	publish,
	R,
	REL,
	remove,
	set,
	targetPart,
	text,
	unset,
	validate,
	W,
	xml,
} from "./docx-package.js";

const object = { additionalProperties: false, minProperties: 1 };
const twips = () => Type.Integer({ description: "Twips (1/20 point)", minimum: 0, maximum: 31680 });
const paragraphProperties = {
	clearRunStyle: Type.Optional(
		Type.Boolean({
			description:
				"Remove direct character styles on ordinary text runs and paragraph marks, not math/text boxes",
		}),
	),
	clearPageBreaks: Type.Optional(
		Type.Boolean({
			description:
				"Remove explicit w:br page breaks in selected paragraphs; leaves line/column breaks",
		}),
	),
	clearStyle: Type.Optional(
		Type.Boolean({ description: "Remove direct paragraph style, not inherited defaults" }),
	),
	font: Type.Optional(Type.String({ minLength: 1, maxLength: 128 })),
	size: Type.Optional(
		Type.Number({
			description: "Font size in points, half-point increments",
			minimum: 1,
			maximum: 400,
		}),
	),
	underline: Type.Optional(Type.Boolean()),
	alignment: Type.Optional(Type.Enum(["left", "center", "right", "both"], { type: "string" })),
	line: Type.Optional(
		Type.Integer({
			description: "Automatic line spacing in 240ths of a line; 360 = 1.5 lines",
			minimum: 1,
			maximum: 2400,
		}),
	),
	before: Type.Optional(twips()),
	after: Type.Optional(twips()),
	left: Type.Optional(twips()),
	hanging: Type.Optional(twips()),
	keepNext: Type.Optional(Type.Boolean()),
	keepLines: Type.Optional(Type.Boolean()),
	pageBreakBefore: Type.Optional(Type.Boolean()),
};
export const docxLayoutSchema = Type.Object(
	{
		part: Type.Optional(
			Type.String({
				description:
					"Story part; default word/document.xml; or word/headerN.xml / word/footerN.xml",
				pattern: "^word/(document|header[0-9]+|footer[0-9]+)\\.xml$",
			}),
		),
		paragraphs: Type.Optional(
			Type.Array(
				Type.Object(
					{
						from: Type.Integer({
							description: "First 1-based paragraph index in the original story",
							minimum: 1,
						}),
						to: Type.Optional(
							Type.Integer({
								description: "Last inclusive paragraph; defaults to from",
								minimum: 1,
							}),
						),
						...paragraphProperties,
					},
					{ additionalProperties: false },
				),
				{ minItems: 1, maxItems: 1000 },
			),
		),
		cloneRows: Type.Optional(
			Type.Array(
				Type.Object(
					{
						table: Type.Integer({
							description: "1-based table index in original story, nested tables excluded",
							minimum: 1,
						}),
						row: Type.Integer({
							description: "1-based source row; clones inserted immediately after it",
							minimum: 1,
						}),
						count: Type.Integer({ minimum: 1, maximum: 100 }),
						content: Type.Enum(["keep", "clear_text"], {
							type: "string",
							description:
								"Explicit content policy; vertical merges and reference-bearing rows rejected",
						}),
						labels: Type.Optional(
							Type.Array(Type.String({ maxLength: 128 }), {
								minItems: 1,
								maxItems: 100,
								description: "One first-cell label per new row; must match count",
							}),
						),
					},
					{ additionalProperties: false },
				),
				{ minItems: 1, maxItems: 100 },
			),
		),
		imageBounds: Type.Optional(
			Type.Object(
				{
					width: Type.Number({
						description: "Maximum width in cm; no upscaling",
						exclusiveMinimum: 0,
						maximum: 100,
					}),
					height: Type.Number({
						description: "Maximum height in cm; aspect ratio preserved",
						exclusiveMinimum: 0,
						maximum: 100,
					}),
				},
				{ additionalProperties: false },
			),
		),
	},
	{
		...object,
		description:
			"Bounded DOCX story layout edits; source indexes from view=structure. Paragraph rules run before row clones. No effective-style or pagination guarantee",
	},
);
export type DocxLayout = Static<typeof docxLayoutSchema>;
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

// Story text only: never traverse math, drawings, or embedded text boxes.
function storyNodes(root: Element): Element[] {
	return elements(root.children, root.ns).flatMap((node) => {
		if (
			["drawing", "pict", "txbxContent", "object"].some((name) => is(node, W, name)) ||
			(node.name.split(":")[0] && node.ns[node.name.split(":")[0]]?.includes("/math"))
		)
			return [];
		return [node, ...storyNodes(node)];
	});
}
function properties(parent: Element, name: string): Element {
	const existing = children(parent, W, name);
	if (existing.length > 1) throw new Error("Ambiguous duplicate DOCX properties");
	if (existing[0]) return existing[0];
	const node = make(parent, W, name);
	parent.children.unshift(node.entry);
	return node;
}
// CT_PPr / CT_RPr order for the properties that this operation writes.
const propertyOrder = [
	"pStyle",
	"rStyle",
	"rFonts",
	"keepNext",
	"keepLines",
	"pageBreakBefore",
	"framePr",
	"widowControl",
	"numPr",
	"suppressLineNumbers",
	"pBdr",
	"shd",
	"tabs",
	"suppressAutoHyphens",
	"kinsoku",
	"wordWrap",
	"overflowPunct",
	"topLinePunct",
	"autoSpaceDE",
	"autoSpaceDN",
	"bidi",
	"adjustRightInd",
	"snapToGrid",
	"spacing",
	"ind",
	"contextualSpacing",
	"mirrorIndents",
	"suppressOverlap",
	"jc",
	"textDirection",
	"textAlignment",
	"textboxTightWrap",
	"outlineLvl",
	"divId",
	"cnfStyle",
	"rPr",
	"sectPr",
	"pPrChange",
];
const runOrder = [
	"rStyle",
	"rFonts",
	"b",
	"bCs",
	"i",
	"iCs",
	"caps",
	"smallCaps",
	"strike",
	"dstrike",
	"outline",
	"shadow",
	"emboss",
	"imprint",
	"noProof",
	"snapToGrid",
	"vanish",
	"webHidden",
	"color",
	"spacing",
	"w",
	"kern",
	"position",
	"sz",
	"szCs",
	"highlight",
	"u",
	"effect",
	"bdr",
	"shd",
	"fitText",
	"vertAlign",
	"rtl",
	"cs",
	"em",
	"lang",
	"eastAsianLayout",
	"specVanish",
	"oMath",
	"rPrChange",
];
function prop(parent: Element, name: string, attrs: Record<string, string | number>) {
	remove(parent, W, name);
	const node = make(parent, W, name);
	for (const [key, value] of Object.entries(attrs)) set(node, W, key, value);
	const order = is(parent, W, "rPr") ? runOrder : propertyOrder;
	const rank = order.indexOf(name);
	const next = elements(parent.children, parent.ns).find((child) => {
		const other = order.indexOf(child.name.split(":").at(-1) ?? "");
		return other >= 0 && other > rank;
	});
	parent.children.splice(
		next ? parent.children.indexOf(next.entry) : parent.children.length,
		0,
		node.entry,
	);
	return node;
}
function update(parent: Element, name: string, values: Record<string, string | number>) {
	const current = children(parent, W, name);
	if (current.length > 1) throw new Error("Ambiguous duplicate DOCX properties");
	const node = current[0] ?? prop(parent, name, {});
	for (const [key, value] of Object.entries(values)) set(node, W, key, value);
	return node;
}
function plainAttrs(node?: Element): Record<string, string> {
	return Object.fromEntries(
		Object.entries(node?.attrs ?? {})
			.filter(([key]) => !key.startsWith("@_xmlns"))
			.slice(0, 8)
			.map(([key, value]) => [key.slice(2, 66), value.slice(0, 128)]),
	);
}
function explicit(parent: Element, name: string) {
	const props = children(parent, W, name)[0];
	return props
		? Object.fromEntries(
				elements(props.children, props.ns)
					.slice(0, 16)
					.map((node) => [node.name.split(":").at(-1), plainAttrs(node)]),
			)
		: {};
}
export async function inspectDocxLayout(
	path: string,
	offset = 0,
	limit = 50,
	selectedPart = "word/document.xml",
) {
	if (
		!Number.isInteger(offset) ||
		offset < 0 ||
		!Number.isInteger(limit) ||
		limit < 1 ||
		limit > 100
	)
		throw new Error("Invalid structure offset/limit (limit 1–100)");
	const { pkg, sha256 } = await load(path);
	const stories = Object.keys(pkg).filter((name) =>
		/^word\/(document|header\d+|footer\d+)\.xml$/.test(name),
	);
	if (!stories.includes(selectedPart)) throw new Error("Unknown structure story part");
	const parts = [selectedPart].map((part) => {
		const root = parse(pkg[part], part);
		const nodes = storyNodes(root);
		const paragraphs = nodes.filter((node) => is(node, W, "p"));
		// Exclude nested tables to keep table coordinates stable and unambiguous.
		const tables = nodes.filter(
			(node) =>
				is(node, W, "tbl") &&
				!nodes.some(
					(ancestor) =>
						is(ancestor, W, "tbl") && all(ancestor).some((child) => child.entry === node.entry),
				),
		);
		let consumedBytes = 0;
		const selected = paragraphs
			.slice(offset, offset + limit)
			.map((p, index) => ({
				index: offset + index + 1,
				text: text(p).slice(0, 100),
				textTruncated: text(p).length > 100,
				runsTruncated: storyNodes(p).filter((node) => is(node, W, "r")).length > 8,
				propertiesTruncated: false,
				pageBreaks: storyNodes(p).filter(
					(node) => is(node, W, "br") && get(node, W, "type") === "page",
				).length,
				explicit: explicit(p, "pPr"),
				runs: storyNodes(p)
					.filter((node) => is(node, W, "r"))
					.slice(0, 8)
					.map((r) => ({ explicit: explicit(r, "rPr") })),
			}))
			.map((record) => {
				if (Buffer.byteLength(JSON.stringify(record)) > 6000) {
					record.runs = [];
					record.runsTruncated = true;
				}
				if (Buffer.byteLength(JSON.stringify(record)) > 6000) {
					record.explicit = {};
					record.propertiesTruncated = true;
				}
				return record;
			})
			.filter((record) => {
				consumedBytes += Buffer.byteLength(JSON.stringify(record));
				return consumedBytes <= 18000;
			});
		return {
			part,
			paragraphCount: paragraphs.length,
			paragraphs: selected,
			nextOffset: offset + selected.length < paragraphs.length ? offset + selected.length : null,
			tableCount: tables.length,
			tables: tables.slice(0, 100).map((table, index) => ({
				index: index + 1,
				rows: children(table, W, "tr").length,
				verticalMerges: all(table).some((node) => is(node, W, "vMerge")),
			})),
		};
	});
	const rels = pkg["word/_rels/document.xml.rels"]
		? children(parse(pkg["word/_rels/document.xml.rels"], "relationships"), REL, "Relationship")
		: [];
	const styleRels = rels.filter((node) => get(node, "", "Type") === `${R}/styles`);
	if (styleRels.length > 1) throw new Error("Ambiguous style relationships");
	const stylePart =
		styleRels[0] && get(styleRels[0], "", "TargetMode") !== "External"
			? targetPart("word", get(styleRels[0], "", "Target") ?? "")
			: undefined;
	const styles = stylePart ? parse(pkg[stylePart], stylePart) : undefined;
	const defaults = styles
		? children(styles, W, "docDefaults")
				.slice(0, 1)
				.map((node) => ({
					run: children(node, W, "rPrDefault")
						.slice(0, 1)
						.map((value) => explicit(value, "rPr")),
					paragraph: children(node, W, "pPrDefault")
						.slice(0, 1)
						.map((value) => explicit(value, "pPr")),
				}))
		: [];
	const document = parse(pkg["word/document.xml"], "word/document.xml");
	const sectionReferences = all(document)
		.filter((node) => is(node, W, "sectPr"))
		.slice(0, 32)
		.flatMap((section, index) =>
			elements(section.children, section.ns)
				.filter((node) => is(node, W, "headerReference") || is(node, W, "footerReference"))
				.slice(0, 6)
				.map((node) => {
					const rel = rels.find((rel) => get(rel, "", "Id") === get(node, R, "id"));
					return {
						section: index + 1,
						kind: is(node, W, "headerReference") ? "header" : "footer",
						type: get(node, W, "type")?.slice(0, 128),
						part: rel ? targetPart("word", get(rel, "", "Target") ?? "").slice(0, 256) : null,
					};
				}),
		);
	return {
		sha256,
		offset,
		sectionReferences,
		stories: stories.slice(0, 100).map((name) => name.slice(0, 256)),
		storiesTruncated: stories.length > 100,
		parts,
		defaults:
			Buffer.byteLength(JSON.stringify(defaults)) > 4000
				? [{ inspectionTruncated: true }]
				: defaults,
		inheritance:
			"Explicit properties only. Absent values inherit through styles, defaults, themes and application settings; effective appearance is not resolved. Run and text excerpts are bounded; pagination requires rendering.",
	};
}
function assertEditable(root: Element) {
	if (
		all(root).some((node) =>
			["ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange", "sdt"].some((name) =>
				is(node, W, name),
			),
		)
	)
		throw new Error("Layout editing revisions or content controls is unsupported");
}
export async function patchDocxLayout(
	inputPath: string,
	outputPath: string,
	spec: DocxLayout,
	dryRun = false,
	overwriteOriginal = false,
) {
	if (!Check(docxLayoutSchema, spec) || (!spec.paragraphs && !spec.cloneRows && !spec.imageBounds))
		throw new Error("Invalid or empty docx_layout");
	const input = resolve(inputPath),
		output = resolve(outputPath);
	if (input !== output || !overwriteOriginal) await distinct(output, [input]);
	const { pkg, sha256 } = await load(input);
	const part = spec.part ?? "word/document.xml";
	const root = parse(pkg[part], part);
	assertEditable(root);
	const nodes = storyNodes(root);
	const paragraphs = nodes.filter((node) => is(node, W, "p"));
	const tables = nodes.filter(
		(node) =>
			is(node, W, "tbl") &&
			!nodes.some(
				(ancestor) =>
					is(ancestor, W, "tbl") && all(ancestor).some((child) => child.entry === node.entry),
			),
	);
	let editedParagraphs = 0,
		clonedRows = 0,
		resizedImages = 0;
	for (const rule of spec.paragraphs ?? []) {
		const end = rule.to ?? rule.from;
		if (
			rule.from > end ||
			end > paragraphs.length ||
			(rule.size !== undefined && !Number.isInteger(rule.size * 2)) ||
			Object.keys(rule).every((key) => key === "from" || key === "to")
		)
			throw new Error("Invalid paragraph range or font size/rule");
		for (const p of paragraphs.slice(rule.from - 1, end)) {
			if (rule.font !== undefined) assertXmlText(rule.font);
			const pp = properties(p, "pPr");
			if (rule.clearStyle) remove(pp, W, "pStyle");
			if (rule.clearPageBreaks)
				for (const parent of [p, ...storyNodes(p)]) {
					const breaks = new Set(
						children(parent, W, "br")
							.filter((node) => get(node, W, "type") === "page")
							.map((node) => node.entry),
					);
					parent.children.splice(
						0,
						parent.children.length,
						...parent.children.filter((entry) => !breaks.has(entry)),
					);
				}
			if (rule.alignment !== undefined) prop(pp, "jc", { val: rule.alignment });
			for (const name of ["keepNext", "keepLines", "pageBreakBefore"] as const)
				if (rule[name] !== undefined) prop(pp, name, { val: rule[name] ? "1" : "0" });
			for (const [name, mapping] of [
				["spacing", { line: rule.line, before: rule.before, after: rule.after }],
				["ind", { left: rule.left, hanging: rule.hanging }],
			] as const) {
				const values = Object.fromEntries(
					Object.entries(mapping).filter(([, value]) => value !== undefined),
				) as Record<string, number | string>;
				if (name === "spacing" && rule.line !== undefined) values.lineRule = "auto";
				if (Object.keys(values).length) {
					const updated = update(pp, name, values);
					// Conflicting automatic/character-unit settings would defeat the requested value.
					const removeNames =
						name === "spacing"
							? [
									rule.before !== undefined && "beforeAutospacing",
									rule.before !== undefined && "beforeLines",
									rule.after !== undefined && "afterAutospacing",
									rule.after !== undefined && "afterLines",
								]
							: [
									rule.hanging !== undefined && "firstLine",
									rule.hanging !== undefined && "firstLineChars",
									rule.hanging !== undefined && "hangingChars",
									rule.left !== undefined && "leftChars",
									rule.left !== undefined && "start",
									rule.left !== undefined && "startChars",
								];
					for (const name of removeNames) if (typeof name === "string") unset(updated, W, name);
				}
			}
			const textRuns = storyNodes(p).filter(
				(node) => is(node, W, "r") && children(node, W, "t").length,
			);
			if (rule.clearRunStyle)
				for (const parent of [pp, ...textRuns]) {
					const rp = children(parent, W, "rPr")[0];
					if (rp) remove(rp, W, "rStyle");
				}
			const typographyTargets: Element[] = [];
			if (rule.font !== undefined || rule.size !== undefined || rule.underline !== undefined) {
				typographyTargets.push(children(pp, W, "rPr")[0] ?? prop(pp, "rPr", {}));
				for (const run of textRuns) typographyTargets.push(properties(run, "rPr"));
			}
			for (const rp of typographyTargets) {
				if (rule.font !== undefined)
					prop(rp, "rFonts", {
						ascii: rule.font,
						hAnsi: rule.font,
						cs: rule.font,
						eastAsia: rule.font,
					});
				if (rule.size !== undefined) {
					prop(rp, "sz", { val: rule.size * 2 });
					prop(rp, "szCs", { val: rule.size * 2 });
				}
				if (rule.underline !== undefined)
					prop(rp, "u", { val: rule.underline ? "single" : "none" });
			}
			editedParagraphs++;
		}
	}
	const originalRows = new Map(tables.map((table) => [table.entry, children(table, W, "tr")]));
	const rowTargets = new Set<string>();
	for (const request of spec.cloneRows ?? []) {
		const key = `${request.table}:${request.row}`;
		if (rowTargets.has(key)) throw new Error("Duplicate row clone target");
		rowTargets.add(key);
		const table = tables[request.table - 1];
		if (!table) throw new Error("Table index out of range");
		if (all(table).some((node) => is(node, W, "vMerge") || is(node, W, "hMerge")))
			throw new Error("Row cloning legacy/vertical merges is unsupported; gridSpan is preserved");
		const row = originalRows.get(table.entry)?.[request.row - 1];
		if (!row || (request.labels && request.labels.length !== request.count))
			throw new Error("Invalid row index or label count");
		if (
			[row, ...all(row)].some(
				(node) =>
					Object.keys(node.attrs).some((name) => {
						const local = name.slice(2).split(":").at(-1) ?? "";
						return (
							get(node, R, local) !== undefined ||
							(/id$/i.test(local) && !["paraId", "textId"].includes(local))
						);
					}) ||
					[
						"drawing",
						"pict",
						"object",
						"bookmarkStart",
						"bookmarkEnd",
						"commentRangeStart",
						"commentRangeEnd",
						"commentReference",
						"fldChar",
						"instrText",
						"fldSimple",
						"footnoteReference",
						"endnoteReference",
						"sectPr",
						"tbl",
						"sdt",
					].some((name) => is(node, W, name)),
			)
		)
			throw new Error("Cannot clone a row with references, fields, nested tables or unique IDs");
		const clones = [];
		for (let i = 0; i < request.count; i++) {
			const clone = elements([structuredClone(row.entry)], table.ns)[0];
			for (const node of [clone, ...all(clone)]) {
				for (const attr of Object.keys(node.attrs))
					if (/:(paraId|textId)$/.test(attr)) delete node.attrs[attr];
				if (request.content === "clear_text" && is(node, W, "t"))
					node.children.splice(0, node.children.length, { "#text": "" });
			}
			if (request.labels) {
				const cell = children(clone, W, "tc")[0];
				const texts = cell ? all(cell).filter((node) => is(node, W, "t")) : [];
				if (!texts.length) throw new Error("First cell has no text for a label");
				texts.forEach((node, index) => {
					assertXmlText(request.labels?.[i] ?? "");
					set(node, "http://www.w3.org/XML/1998/namespace", "space", "preserve");
					node.children.splice(0, node.children.length, {
						"#text": index === 0 ? (request.labels?.[i] ?? "") : "",
					});
				});
			}
			clones.push(clone.entry);
		}
		table.children.splice(table.children.indexOf(row.entry) + 1, 0, ...clones);
		clonedRows += clones.length;
	}
	if (spec.imageBounds)
		for (const drawing of nodes
			.filter((node) => is(node, W, "p"))
			.flatMap((p) => all(p).filter((node) => is(node, W, "drawing")))) {
			if (!all(drawing).some((node) => is(node, A, "blip"))) continue;
			const extents = all(drawing).filter((node) => is(node, WP, "extent"));
			if (extents.length !== 1) throw new Error("Unsupported drawing extent geometry");
			const cx = Number(get(extents[0], "", "cx")),
				cy = Number(get(extents[0], "", "cy"));
			if (!(cx > 0 && cy > 0 && Number.isFinite(cx) && Number.isFinite(cy)))
				throw new Error("Invalid image dimensions");
			const scale = Math.min(
				1,
				(spec.imageBounds.width * 360000) / cx,
				(spec.imageBounds.height * 360000) / cy,
			);
			if (scale === 1) continue;
			for (const extent of all(drawing).filter(
				(node) => is(node, WP, "extent") || is(node, A, "ext"),
			)) {
				for (const dimension of ["cx", "cy"]) {
					const value = get(extent, "", dimension);
					if (value !== undefined) {
						if (!Number.isSafeInteger(Number(value)) || Number(value) < 0)
							throw new Error("Invalid drawing transform dimensions");
						set(extent, "", dimension, Math.max(1, Math.round(Number(value) * scale)));
					}
				}
			}
			resizedImages++;
		}
	pkg[part] = xml(root);
	validate(pkg);
	if (!dryRun) await publish(pkg, output);
	return {
		sha256,
		changedParts: [part],
		editedParagraphs,
		clonedRows,
		resizedImages,
		written: !dryRun,
		warnings: [
			"Explicit layout edits only; render to check pagination. Math and text boxes are not formatted; legacy VML images are not resized.",
		],
	};
}
