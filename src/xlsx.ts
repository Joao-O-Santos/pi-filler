import { randomUUID } from "node:crypto";
import { mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import { dirname, join, posix, resolve } from "node:path";
import { parse as parseCsv } from "csv-parse/sync";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { type OfficePackage, packOfficePackage, unpackOfficePackage } from "./ooxml.js";

const MAX_COLUMN = 16_384;
const MAX_ROW = 1_048_576;
const MAX_MUTATED_CELLS = 1_000_000;

type Package = OfficePackage;
type XmlObject = Record<string, unknown>;

export interface CellAddress {
	column: number;
	row: number;
}

export interface XlsxSheetStructure {
	sheet: string;
	usedRange: string | null;
	rows: number;
	columns: number;
	formulaCells: number;
	mergedRanges: number;
	styledCells: number;
}

export interface XlsxStructure {
	sheets: string[];
	activeSheet: string | null;
	worksheets: XlsxSheetStructure[];
}

export interface FillOptions {
	sheet: string;
	startCell: string;
	hasHeader?: boolean;
	valueMode?: "text" | "auto";
	dryRun?: boolean;
	signal?: AbortSignal;
}

export interface XlsxWriteResult {
	sheet: string;
	targetRange: string;
	writtenRange: string;
	rows: number;
	columns: number;
	formulaConflicts: number;
	changedParts: string[];
	output: string;
	wouldWrite: boolean;
	written: boolean;
}

export interface XlsxPatchSet {
	font?: { bold?: boolean; italic?: boolean; size?: number };
	fill?: string;
	alignment?: {
		horizontal?: "left" | "center" | "right";
		vertical?: "top" | "center" | "bottom";
		wrap_text?: boolean;
	};
	number_format?: string;
	border?: { style?: "thin" | "medium" | "thick" };
	column_width?: number;
	row_height?: number;
}

export interface XlsxPatch {
	sheet: string;
	range: string;
	patches: XlsxPatchSet;
}

export interface XlsxPatchResult {
	sheet: string;
	range: string;
	styledCells: number;
	columnsChanged: number;
	rowsChanged: number;
	changedParts: string[];
	output: string;
	wouldWrite: boolean;
	written: boolean;
}

interface ResolvedSheet {
	name: string;
	part: string;
	worksheet: boolean;
}

interface CellRange {
	start: CellAddress;
	end: CellAddress;
}

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	parseTagValue: false,
	parseAttributeValue: false,
	trimValues: false,
});
const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: "@_",
	format: false,
	suppressBooleanAttributes: false,
	suppressEmptyNode: false,
});

function isObject(value: unknown): value is XmlObject {
	return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function localName(name: string): string {
	return name.replace(/^@_/, "").split(":").at(-1) ?? name;
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

function setAttr(value: XmlObject, name: string, setting: string | number): void {
	const existing = Object.keys(value).find(
		(key) => key.startsWith("@_") && localName(key) === name,
	);
	value[existing ?? `@_${name}`] = String(setting);
}

function removeAttr(value: XmlObject, name: string): void {
	for (const key of Object.keys(value)) {
		if (key.startsWith("@_") && localName(key) === name) delete value[key];
	}
}

function directKey(value: XmlObject | undefined, name: string): string | undefined {
	if (!value) return undefined;
	return Object.keys(value).find((key) => !key.startsWith("@_") && localName(key) === name);
}

function directObject(value: XmlObject | undefined, name: string): XmlObject | undefined {
	const key = directKey(value, name);
	if (!key || !value) return undefined;
	const child = value[key];
	const candidate = Array.isArray(child) ? child[0] : child;
	return isObject(candidate) ? candidate : undefined;
}

function directObjects(value: XmlObject | undefined, name: string): XmlObject[] {
	const key = directKey(value, name);
	if (!key || !value) return [];
	const child = value[key];
	return (Array.isArray(child) ? child : [child]).filter(isObject);
}

function elementName(parent: XmlObject, name: string, fallbackPrefix?: string): string {
	const existing = directKey(parent, name);
	if (existing) return existing;
	const sibling = Object.keys(parent).find((key) => !key.startsWith("@_") && key.includes(":"));
	const prefix = sibling?.split(":", 1)[0] ?? fallbackPrefix;
	return prefix ? `${prefix}:${name}` : name;
}

function xmlPrefix(value: XmlObject, name: string): string | undefined {
	const key = Object.keys(value).find(
		(candidate) => !candidate.startsWith("@_") && localName(candidate) === name,
	);
	return key?.includes(":") ? key.split(":", 1)[0] : undefined;
}

function setDirectObjects(
	parent: XmlObject,
	name: string,
	values: XmlObject[],
	fallbackPrefix?: string,
): void {
	const key = directKey(parent, name) ?? elementName(parent, name, fallbackPrefix);
	parent[key] = values;
}

function insertDirectBefore(
	parent: XmlObject,
	name: string,
	value: unknown,
	beforeName: string,
	fallbackPrefix?: string,
): void {
	const key = elementName(parent, name, fallbackPrefix);
	const entries = Object.entries(parent).filter(([entry]) => entry !== key);
	for (const existing of Object.keys(parent)) delete parent[existing];
	let inserted = false;
	for (const [entry, child] of entries) {
		if (!inserted && !entry.startsWith("@_") && localName(entry) === beforeName) {
			parent[key] = value;
			inserted = true;
		}
		parent[entry] = child;
	}
	if (!inserted) parent[key] = value;
}

function removeDirect(parent: XmlObject, ...names: string[]): void {
	for (const key of Object.keys(parent)) {
		if (!key.startsWith("@_") && names.includes(localName(key))) delete parent[key];
	}
}

function rootObject(value: XmlObject, name: string, _part: string): XmlObject {
	const root = directObject(value, name);
	if (!root) throw new Error(`XLSX package is missing a required ${name} element`);
	return root;
}

function parseXml(bytes: Uint8Array | undefined, _part: string): XmlObject {
	if (!bytes) throw new Error("XLSX package is missing a required part");
	try {
		return parser.parse(new TextDecoder().decode(bytes), true) as XmlObject;
	} catch {
		throw new Error("Invalid XML in XLSX package part");
	}
}

function serializeXml(value: XmlObject): Uint8Array {
	return new TextEncoder().encode(builder.build(value));
}

function unpack(bytes: Uint8Array): Package {
	return unpackOfficePackage(bytes, "XLSX");
}

async function load(path: string): Promise<Package> {
	return unpack(new Uint8Array(await readFile(resolve(path))));
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
		const relationships = parseXml(pkg[part], part);
		for (const relationship of directObjects(
			rootObject(relationships, "Relationships", part),
			"Relationship",
		)) {
			const target = attr(relationship, "Target");
			if (!target || attr(relationship, "TargetMode") === "External") continue;
			const normalized = relationshipTarget(part, target);
			if (!pkg[normalized]) throw new Error("XLSX package contains a missing relationship target");
		}
	}
}

function workbookSheets(pkg: Package): { sheets: ResolvedSheet[]; activeIndex: number } {
	const workbookPart = "xl/workbook.xml";
	const relationshipsPart = "xl/_rels/workbook.xml.rels";
	const workbook = rootObject(parseXml(pkg[workbookPart], workbookPart), "workbook", workbookPart);
	const relationships = rootObject(
		parseXml(pkg[relationshipsPart], relationshipsPart),
		"Relationships",
		relationshipsPart,
	);
	const targets = new Map<string, { part: string; type: string }>();
	for (const relationship of directObjects(relationships, "Relationship")) {
		const id = attr(relationship, "Id");
		const target = attr(relationship, "Target");
		if (id && target && attr(relationship, "TargetMode") !== "External") {
			targets.set(id, {
				part: relationshipTarget(relationshipsPart, target),
				type: attr(relationship, "Type") ?? "",
			});
		}
	}
	const sheetContainer = directObject(workbook, "sheets");
	if (!sheetContainer) throw new Error("XLSX workbook has no sheets");
	const sheets = directObjects(sheetContainer, "sheet").map((sheet) => {
		const name = attr(sheet, "name");
		const relationshipId = attr(sheet, "id");
		const relationship = relationshipId ? targets.get(relationshipId) : undefined;
		const part = relationship?.part;
		if (!name || !part || !pkg[part]) {
			throw new Error("XLSX workbook contains an unresolvable sheet");
		}
		const worksheet = relationship?.type.endsWith("/worksheet") ?? false;
		return { name, part, worksheet };
	});
	if (sheets.length === 0) throw new Error("XLSX workbook has no sheets");
	const workbookView = directObject(directObject(workbook, "bookViews"), "workbookView");
	const activeIndex = Number(attr(workbookView, "activeTab") ?? "0");
	return { sheets, activeIndex: Number.isInteger(activeIndex) ? activeIndex : 0 };
}

function validatePackage(pkg: Package): void {
	for (const part of [
		"[Content_Types].xml",
		"_rels/.rels",
		"xl/workbook.xml",
		"xl/_rels/workbook.xml.rels",
	]) {
		parseXml(pkg[part], part);
	}
	validateRelationships(pkg);
	const { sheets } = workbookSheets(pkg);
	for (const sheet of sheets.filter((candidate) => candidate.worksheet)) {
		rootObject(parseXml(pkg[sheet.part], sheet.part), "worksheet", sheet.part);
	}
	if (pkg["xl/styles.xml"]) {
		rootObject(parseXml(pkg["xl/styles.xml"], "xl/styles.xml"), "styleSheet", "xl/styles.xml");
	}
}

function resolveSheet(pkg: Package, name: string): ResolvedSheet {
	const sheet = workbookSheets(pkg).sheets.find((candidate) => candidate.name === name);
	if (!sheet) throw new Error("Requested XLSX sheet does not exist");
	if (!sheet.worksheet) throw new Error("Requested XLSX sheet is not a worksheet");
	return sheet;
}

export function columnNameToNumber(name: string): number {
	if (!/^[A-Za-z]+$/.test(name)) throw new Error("Invalid XLSX column name");
	let column = 0;
	for (const character of name.toUpperCase()) {
		column = column * 26 + character.charCodeAt(0) - 64;
		if (column > MAX_COLUMN) throw new Error("XLSX column exceeds Excel limits");
	}
	return column;
}

export function columnNumberToName(column: number): string {
	if (!Number.isInteger(column) || column < 1 || column > MAX_COLUMN) {
		throw new Error("XLSX column number is outside Excel limits");
	}
	let current = column;
	let name = "";
	while (current > 0) {
		current -= 1;
		name = String.fromCharCode(65 + (current % 26)) + name;
		current = Math.floor(current / 26);
	}
	return name;
}

export function parseCellAddress(address: string): CellAddress {
	const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(address);
	if (!match) throw new Error("Invalid XLSX cell address");
	const row = Number(match[2]);
	if (!Number.isSafeInteger(row) || row > MAX_ROW) {
		throw new Error("XLSX row is outside Excel limits");
	}
	return { column: columnNameToNumber(match[1]), row };
}

function formatCell(address: CellAddress): string {
	return `${columnNumberToName(address.column)}${address.row}`;
}

export function offsetCell(address: string, rowOffset: number, columnOffset: number): string {
	if (!Number.isInteger(rowOffset) || !Number.isInteger(columnOffset)) {
		throw new Error("XLSX cell offsets must be integers");
	}
	const parsed = parseCellAddress(address);
	const row = parsed.row + rowOffset;
	const column = parsed.column + columnOffset;
	if (row < 1 || row > MAX_ROW || column < 1 || column > MAX_COLUMN) {
		throw new Error("Offset XLSX cell is outside Excel limits");
	}
	return formatCell({ row, column });
}

export function rangeFromDimensions(start: string, rows: number, columns: number): string {
	if (!Number.isInteger(rows) || rows < 1 || !Number.isInteger(columns) || columns < 1) {
		throw new Error("XLSX range dimensions must be positive integers");
	}
	const end = offsetCell(start, rows - 1, columns - 1);
	return `${formatCell(parseCellAddress(start))}:${end}`;
}

function parseRange(range: string): CellRange {
	const parts = range.split(":");
	if (parts.length < 1 || parts.length > 2) throw new Error("Invalid XLSX range");
	const start = parseCellAddress(parts[0]);
	const end = parseCellAddress(parts[1] ?? parts[0]);
	if (start.row > end.row || start.column > end.column) {
		throw new Error("XLSX range must run from upper-left to lower-right");
	}
	return { start, end };
}

function cellCount(range: CellRange): number {
	return (range.end.row - range.start.row + 1) * (range.end.column - range.start.column + 1);
}

function sheetData(worksheet: XmlObject, prefix?: string): XmlObject {
	let data = directObject(worksheet, "sheetData");
	if (!data) {
		data = {};
		worksheet[elementName(worksheet, "sheetData", prefix)] = data;
	}
	return data;
}

function rowsByNumber(worksheet: XmlObject, prefix?: string): Map<number, XmlObject> {
	const result = new Map<number, XmlObject>();
	for (const row of directObjects(sheetData(worksheet, prefix), "row")) {
		const number = Number(attr(row, "r"));
		if (Number.isInteger(number) && number >= 1 && number <= MAX_ROW) result.set(number, row);
	}
	return result;
}

function effectiveStyleIndex(
	worksheet: XmlObject,
	row: XmlObject,
	cell: XmlObject | undefined,
	column: number,
): number {
	if (cell && attr(cell, "s") !== undefined) return Number(attr(cell, "s"));
	if (["1", "true"].includes(attr(row, "customFormat") ?? "") && attr(row, "s") !== undefined)
		return Number(attr(row, "s"));
	const definition = directObjects(directObject(worksheet, "cols"), "col")
		.filter(
			(candidate) =>
				Number(attr(candidate, "min")) <= column && Number(attr(candidate, "max")) >= column,
		)
		.at(-1);
	return Number(attr(definition, "style") ?? "0");
}

function cellsByAddress(row: XmlObject): Map<string, XmlObject> {
	const result = new Map<string, XmlObject>();
	for (const cell of directObjects(row, "c")) {
		const reference = attr(cell, "r")?.toUpperCase();
		if (reference) result.set(reference, cell);
	}
	return result;
}

function ensureRow(
	_worksheet: XmlObject,
	rows: Map<number, XmlObject>,
	rowNumber: number,
): XmlObject {
	let row = rows.get(rowNumber);
	if (!row) {
		row = { "@_r": String(rowNumber) };
		rows.set(rowNumber, row);
	}
	return row;
}

function syncRows(worksheet: XmlObject, rows: Map<number, XmlObject>, prefix?: string): void {
	setDirectObjects(
		sheetData(worksheet, prefix),
		"row",
		[...rows.entries()].sort(([left], [right]) => left - right).map((entry) => entry[1]),
		prefix,
	);
}

function ensureCell(cells: Map<string, XmlObject>, address: string): XmlObject {
	let cell = cells.get(address);
	if (!cell) {
		cell = { "@_r": address };
		cells.set(address, cell);
	}
	return cell;
}

function syncCells(row: XmlObject, cells: Map<string, XmlObject>, prefix?: string): void {
	setDirectObjects(
		row,
		"c",
		[...cells.values()].sort((left, right) => {
			return (
				parseCellAddress(attr(left, "r") ?? "A1").column -
				parseCellAddress(attr(right, "r") ?? "A1").column
			);
		}),
		prefix,
	);
}

function hasFormula(cell: XmlObject): boolean {
	return directKey(cell, "f") !== undefined;
}

function formulaRanges(worksheet: XmlObject): CellRange[] {
	const ranges: CellRange[] = [];
	for (const row of directObjects(sheetData(worksheet), "row")) {
		for (const cell of directObjects(row, "c")) {
			const formulaKey = directKey(cell, "f");
			if (!formulaKey) continue;
			const formula = cell[formulaKey];
			const reference = (isObject(formula) ? attr(formula, "ref") : undefined) ?? attr(cell, "r");
			if (!reference) continue;
			try {
				ranges.push(parseRange(reference));
			} catch {
				throw new Error("XLSX worksheet contains an invalid formula range");
			}
		}
	}
	return ranges;
}

function overlaps(left: CellRange, right: CellRange): boolean {
	return (
		left.start.row <= right.end.row &&
		right.start.row <= left.end.row &&
		left.start.column <= right.end.column &&
		right.start.column <= left.end.column
	);
}

function setCellValue(
	cell: XmlObject,
	value: string,
	mode: "text" | "auto",
	prefix?: string,
): void {
	removeDirect(cell, "f", "v", "is");
	removeAttr(cell, "t");
	if (mode === "auto" && value === "") return;
	if (mode === "auto" && /^(?:true|false)$/i.test(value)) {
		setAttr(cell, "t", "b");
		insertDirectBefore(cell, "v", value.toLowerCase() === "true" ? "1" : "0", "extLst", prefix);
		return;
	}
	const numeric = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(value);
	const number = numeric ? Number(value) : Number.NaN;
	const safeInteger = !/^[-+]?\d+$/.test(value) || Number.isSafeInteger(number);
	if (mode === "auto" && numeric && Number.isFinite(number) && safeInteger) {
		insertDirectBefore(cell, "v", value, "extLst", prefix);
		return;
	}
	setAttr(cell, "t", "inlineStr");
	const text: XmlObject = { "#text": value, "@_xml:space": "preserve" };
	insertDirectBefore(cell, "is", { [elementName(text, "t", prefix)]: text }, "extLst", prefix);
}

function cellMatchesValue(
	cell: XmlObject | undefined,
	value: string,
	mode: "text" | "auto",
): boolean {
	if (mode === "auto" && value === "") return !cell || !directKey(cell, "v");
	if (mode === "auto" && /^(?:true|false)$/i.test(value)) {
		return (
			attr(cell, "t") === "b" &&
			String(cell?.[directKey(cell, "v") ?? ""]) === (value.toLowerCase() === "true" ? "1" : "0")
		);
	}
	const numeric = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(value);
	const number = numeric ? Number(value) : Number.NaN;
	const safeInteger = !/^[-+]?\d+$/.test(value) || Number.isSafeInteger(number);
	if (mode === "auto" && numeric && Number.isFinite(number) && safeInteger) {
		return attr(cell, "t") === undefined && String(cell?.[directKey(cell, "v") ?? ""]) === value;
	}
	const inline = directObject(cell, "is");
	const textNode = inline ? directKey(inline, "t") : undefined;
	const textValue = textNode ? inline?.[textNode] : undefined;
	const text = isObject(textValue) ? textValue["#text"] : textValue;
	return attr(cell, "t") === "inlineStr" && String(text ?? "") === value;
}

function actualUsedRange(worksheet: XmlObject): CellRange | null {
	let minimumRow = MAX_ROW + 1;
	let minimumColumn = MAX_COLUMN + 1;
	let maximumRow = 0;
	let maximumColumn = 0;
	for (const row of directObjects(sheetData(worksheet), "row")) {
		for (const cell of directObjects(row, "c")) {
			const reference = attr(cell, "r");
			if (!reference) continue;
			try {
				const address = parseCellAddress(reference);
				minimumRow = Math.min(minimumRow, address.row);
				minimumColumn = Math.min(minimumColumn, address.column);
				maximumRow = Math.max(maximumRow, address.row);
				maximumColumn = Math.max(maximumColumn, address.column);
			} catch {
				throw new Error("XLSX worksheet contains an invalid cell reference");
			}
		}
	}
	for (const merge of directObjects(directObject(worksheet, "mergeCells"), "mergeCell")) {
		const reference = attr(merge, "ref");
		if (!reference) continue;
		try {
			const range = parseRange(reference);
			minimumRow = Math.min(minimumRow, range.start.row);
			minimumColumn = Math.min(minimumColumn, range.start.column);
			maximumRow = Math.max(maximumRow, range.end.row);
			maximumColumn = Math.max(maximumColumn, range.end.column);
		} catch {
			throw new Error("XLSX worksheet contains an invalid merged range");
		}
	}
	if (maximumRow === 0) return null;
	return {
		start: { row: minimumRow, column: minimumColumn },
		end: { row: maximumRow, column: maximumColumn },
	};
}

function updateDimension(worksheet: XmlObject, prefix?: string): void {
	const used = actualUsedRange(worksheet);
	let dimension = directObject(worksheet, "dimension");
	if (!used) {
		if (dimension) setAttr(dimension, "ref", "A1");
		return;
	}
	if (!dimension) {
		dimension = {};
		const before = directKey(worksheet, "cols") ? "cols" : "sheetData";
		const key = elementName(worksheet, "dimension", prefix);
		const entries = Object.entries(worksheet).filter(([entry]) => entry !== key);
		for (const entry of Object.keys(worksheet)) delete worksheet[entry];
		let inserted = false;
		for (const [entry, child] of entries) {
			if (!inserted && !entry.startsWith("@_") && localName(entry) === before) {
				worksheet[key] = dimension;
				inserted = true;
			}
			worksheet[entry] = child;
		}
		if (!inserted) worksheet[key] = dimension;
	}
	const start = formatCell(used.start);
	const end = formatCell(used.end);
	setAttr(dimension, "ref", start === end ? start : `${start}:${end}`);
}

function inspectSheet(pkg: Package, sheet: ResolvedSheet): XlsxSheetStructure {
	const worksheet = rootObject(parseXml(pkg[sheet.part], sheet.part), "worksheet", sheet.part);
	const used = actualUsedRange(worksheet);
	let formulaCells = 0;
	let styledCells = 0;
	for (const row of directObjects(sheetData(worksheet), "row")) {
		for (const cell of directObjects(row, "c")) {
			if (hasFormula(cell)) formulaCells += 1;
			if (attr(cell, "s") !== undefined) styledCells += 1;
		}
	}
	return {
		sheet: sheet.name,
		usedRange: used
			? `${formatCell(used.start)}${used.start.row === used.end.row && used.start.column === used.end.column ? "" : `:${formatCell(used.end)}`}`
			: null,
		rows: used ? used.end.row - used.start.row + 1 : 0,
		columns: used ? used.end.column - used.start.column + 1 : 0,
		formulaCells,
		mergedRanges: directObjects(directObject(worksheet, "mergeCells"), "mergeCell").length,
		styledCells,
	};
}

export async function inspectXlsxStructure(path: string): Promise<XlsxStructure> {
	const pkg = await load(path);
	validatePackage(pkg);
	const workbook = workbookSheets(pkg);
	return {
		sheets: workbook.sheets.map((sheet) => sheet.name),
		activeSheet: workbook.sheets[workbook.activeIndex]?.name ?? workbook.sheets[0]?.name ?? null,
		worksheets: workbook.sheets
			.filter((sheet) => sheet.worksheet)
			.map((sheet) => inspectSheet(pkg, sheet)),
	};
}

function validateFilledPackage(
	pkg: Package,
	sheetName: string,
	destination: CellRange,
	rows: string[][],
	mode: "text" | "auto",
): void {
	const sheet = resolveSheet(pkg, sheetName);
	const worksheet = rootObject(parseXml(pkg[sheet.part], sheet.part), "worksheet", sheet.part);
	const used = actualUsedRange(worksheet);
	if (
		!used ||
		used.start.row > destination.start.row ||
		used.start.column > destination.start.column ||
		used.end.row < destination.end.row ||
		used.end.column < destination.end.column
	) {
		throw new Error("XLSX output validation failed for destination range");
	}
	const rowsByNumberMap = rowsByNumber(
		worksheet,
		xmlPrefix(parseXml(pkg[sheet.part], sheet.part), "worksheet"),
	);
	for (let row = destination.start.row; row <= destination.end.row; row += 1) {
		const cells = cellsByAddress(rowsByNumberMap.get(row) ?? {});
		for (let column = destination.start.column; column <= destination.end.column; column += 1) {
			const cell = cells.get(formatCell({ row, column }));
			const expected = rows[row - destination.start.row][column - destination.start.column] ?? "";
			if (!cellMatchesValue(cell, expected, mode)) {
				throw new Error("XLSX output validation failed for destination values");
			}
		}
	}
}

function parseCsvRows(bytes: Uint8Array): string[][] {
	try {
		return parseCsv(new TextDecoder("utf-8", { fatal: true }).decode(bytes), {
			bom: true,
			relax_column_count: true,
			skip_empty_lines: false,
		}) as string[][];
	} catch {
		throw new Error("Invalid UTF-8 or malformed CSV input");
	}
}

function throwIfAborted(signal?: AbortSignal): void {
	if (signal?.aborted) throw new Error("XLSX operation aborted");
}

async function writePackageAtomically(
	pkg: Package,
	output: string,
	validate: (path: string) => Promise<void>,
	signal?: AbortSignal,
): Promise<void> {
	const temporary = join(dirname(output), `.${output.split(/[\\/]/).pop()}.${randomUUID()}.tmp`);
	await mkdir(dirname(output), { recursive: true });
	try {
		await writeFile(temporary, packOfficePackage(pkg));
		await validate(temporary);
		throwIfAborted(signal);
		await rename(temporary, output);
	} finally {
		await rm(temporary, { force: true });
	}
}

export async function fillXlsxFromCsv(
	templatePath: string,
	csvPath: string,
	outputPath: string,
	options: FillOptions,
): Promise<XlsxWriteResult> {
	const template = resolve(templatePath);
	const output = resolve(outputPath);
	if (template === output) throw new Error("XLSX write output must differ from input");
	const start = parseCellAddress(options.startCell);
	throwIfAborted(options.signal);
	const pkg = await load(template);
	validatePackage(pkg);
	const sheet = resolveSheet(pkg, options.sheet);
	const allRows = parseCsvRows(new Uint8Array(await readFile(resolve(csvPath))));
	throwIfAborted(options.signal);
	const rows = (options.hasHeader ?? true) ? allRows.slice(1) : allRows;
	if (rows.length === 0) throw new Error("CSV contains no destination data rows");
	const columns = Math.max(0, ...rows.map((row) => row.length));
	if (columns === 0) throw new Error("CSV contains no destination columns");
	const destination = parseRange(rangeFromDimensions(options.startCell, rows.length, columns));
	if (cellCount(destination) > MAX_MUTATED_CELLS) {
		throw new Error("XLSX destination exceeds the supported mutation limit");
	}
	const targetRange = `${formatCell(destination.start)}:${formatCell(destination.end)}`;
	const worksheetDocument = parseXml(pkg[sheet.part], sheet.part);
	const worksheet = rootObject(worksheetDocument, "worksheet", sheet.part);
	const prefix = xmlPrefix(worksheetDocument, "worksheet");
	const existingRows = rowsByNumber(worksheet, prefix);
	let formulaConflicts = 0;
	for (const formulaRange of formulaRanges(worksheet)) {
		if (overlaps(destination, formulaRange)) formulaConflicts += 1;
	}
	if (formulaConflicts > 0) {
		throw new Error(
			`Destination intersects ${formulaConflicts} formula cell${formulaConflicts === 1 ? "" : "s"}. Refusing to overwrite formulas.`,
		);
	}
	const result: XlsxWriteResult = {
		sheet: sheet.name,
		targetRange,
		writtenRange: targetRange,
		rows: rows.length,
		columns,
		formulaConflicts,
		changedParts: [sheet.part],
		output,
		wouldWrite: true,
		written: false,
	};
	const mode = options.valueMode ?? "text";
	for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
		throwIfAborted(options.signal);
		const row = ensureRow(worksheet, existingRows, start.row + rowOffset);
		const cells = cellsByAddress(row);
		for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
			const address = formatCell({
				row: start.row + rowOffset,
				column: start.column + columnOffset,
			});
			setCellValue(ensureCell(cells, address), rows[rowOffset][columnOffset] ?? "", mode, prefix);
		}
		syncCells(row, cells, prefix);
	}
	syncRows(worksheet, existingRows, prefix);
	updateDimension(worksheet, prefix);
	pkg[sheet.part] = serializeXml(worksheetDocument);
	validatePackage(pkg);
	validateFilledPackage(pkg, sheet.name, destination, rows, mode);
	throwIfAborted(options.signal);
	if (options.dryRun) return result;
	await writePackageAtomically(
		pkg,
		output,
		async (temporary) => {
			const candidate = await load(temporary);
			validatePackage(candidate);
			validateFilledPackage(candidate, sheet.name, destination, rows, mode);
		},
		options.signal,
	);
	result.written = true;
	return result;
}

function stable(value: unknown): string {
	if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
	if (isObject(value)) {
		if (Object.keys(value).length === 0) return JSON.stringify("");
		return `{${Object.keys(value)
			.sort()
			.map((key) => `${JSON.stringify(key)}:${stable(value[key])}`)
			.join(",")}}`;
	}
	return JSON.stringify(value);
}

function clone(value: XmlObject): XmlObject {
	return structuredClone(value);
}

function styleCollection(styles: XmlObject, name: string, itemName: string): XmlObject[] {
	const collection = directObject(styles, name);
	if (!collection) throw new Error(`XLSX styles are missing required ${name}`);
	return directObjects(collection, itemName);
}

function replaceStyleCollection(
	styles: XmlObject,
	name: string,
	itemName: string,
	items: XmlObject[],
): void {
	const collection = directObject(styles, name);
	if (!collection) throw new Error(`XLSX styles are missing required ${name}`);
	setDirectObjects(collection, itemName, items);
	setAttr(collection, "count", items.length);
}

function componentIndex(items: XmlObject[], requested: XmlObject): number {
	const representation = stable(requested);
	const found = items.findIndex((item) => stable(item) === representation);
	if (found >= 0) return found;
	items.push(requested);
	return items.length - 1;
}

function setBooleanChild(value: XmlObject, name: string, setting: boolean, prefix?: string): void {
	removeDirect(value, name);
	if (setting) value[elementName(value, name, prefix)] = {};
}

function setValueChild(value: XmlObject, name: string, setting: number, prefix?: string): void {
	removeDirect(value, name);
	value[elementName(value, name, prefix)] = { "@_val": String(setting) };
}

function q(prefix: string | undefined, name: string): string {
	return prefix ? `${prefix}:${name}` : name;
}

function styleIndexFor(styles: XmlObject, baseIndex: number, patch: XlsxPatchSet): number {
	const prefix =
		xmlPrefix(styles, "fonts") ?? xmlPrefix(styles, "fills") ?? xmlPrefix(styles, "cellXfs");
	const fonts = styleCollection(styles, "fonts", "font");
	const fills = styleCollection(styles, "fills", "fill");
	const borders = styleCollection(styles, "borders", "border");
	const cellXfs = styleCollection(styles, "cellXfs", "xf");
	if (!cellXfs[baseIndex]) throw new Error("XLSX worksheet contains an invalid style index");
	const base = clone(cellXfs[baseIndex]);

	if (patch.font) {
		const fontId = Number(attr(base, "fontId") ?? "0");
		const font = clone(fonts[fontId] ?? fonts[0] ?? {});
		if (patch.font.bold !== undefined) setBooleanChild(font, "b", patch.font.bold, prefix);
		if (patch.font.italic !== undefined) setBooleanChild(font, "i", patch.font.italic, prefix);
		if (patch.font.size !== undefined) setValueChild(font, "sz", patch.font.size, prefix);
		setAttr(base, "fontId", componentIndex(fonts, font));
		setAttr(base, "applyFont", "1");
	}
	if (patch.fill !== undefined) {
		const fill: XmlObject = {
			[q(prefix, "patternFill")]: {
				"@_patternType": "solid",
				[q(prefix, "fgColor")]: { "@_rgb": `FF${patch.fill.toUpperCase()}` },
				[q(prefix, "bgColor")]: { "@_indexed": "64" },
			},
		};
		setAttr(base, "fillId", componentIndex(fills, fill));
		setAttr(base, "applyFill", "1");
	}
	if (patch.border) {
		const edge = patch.border.style ? { "@_style": patch.border.style } : {};
		const border = {
			[q(prefix, "left")]: clone(edge),
			[q(prefix, "right")]: clone(edge),
			[q(prefix, "top")]: clone(edge),
			[q(prefix, "bottom")]: clone(edge),
		};
		setAttr(base, "borderId", componentIndex(borders, border));
		setAttr(base, "applyBorder", "1");
	}
	if (patch.alignment) {
		const alignment = clone(directObject(base, "alignment") ?? {});
		if (patch.alignment.horizontal !== undefined) {
			setAttr(alignment, "horizontal", patch.alignment.horizontal);
		}
		if (patch.alignment.vertical !== undefined) {
			setAttr(alignment, "vertical", patch.alignment.vertical);
		}
		if (patch.alignment.wrap_text !== undefined) {
			if (patch.alignment.wrap_text) setAttr(alignment, "wrapText", "1");
			else removeAttr(alignment, "wrapText");
		}
		removeDirect(base, "alignment");
		base[elementName(base, "alignment", prefix)] = alignment;
		setAttr(base, "applyAlignment", "1");
	}
	if (patch.number_format !== undefined) {
		let numberFormats = directObject(styles, "numFmts");
		if (!numberFormats) {
			numberFormats = {};
			insertDirectBefore(styles, "numFmts", numberFormats, "fonts");
		}
		const formats = directObjects(numberFormats, "numFmt");
		let format = formats.find((candidate) => attr(candidate, "formatCode") === patch.number_format);
		if (!format) {
			const used = formats.map((candidate) => Number(attr(candidate, "numFmtId") ?? "0"));
			const id = Math.max(163, ...used.filter(Number.isFinite)) + 1;
			format = { "@_numFmtId": String(id), "@_formatCode": patch.number_format };
			formats.push(format);
			setDirectObjects(numberFormats, "numFmt", formats);
			setAttr(numberFormats, "count", formats.length);
		}
		setAttr(base, "numFmtId", attr(format, "numFmtId") ?? "0");
		setAttr(base, "applyNumberFormat", "1");
	}

	replaceStyleCollection(styles, "fonts", "font", fonts);
	replaceStyleCollection(styles, "fills", "fill", fills);
	replaceStyleCollection(styles, "borders", "border", borders);
	const index = componentIndex(cellXfs, base);
	replaceStyleCollection(styles, "cellXfs", "xf", cellXfs);
	return index;
}

function patchColumns(worksheet: XmlObject, range: CellRange, width: number): number {
	let columns = directObject(worksheet, "cols");
	if (!columns) {
		columns = {};
		insertDirectBefore(worksheet, "cols", columns, "sheetData");
	}
	const existing = directObjects(columns, "col");
	const definitions = existing.map((column) => {
		const minimum = Number(attr(column, "min"));
		const maximum = Number(attr(column, "max"));
		if (!Number.isInteger(minimum) || !Number.isInteger(maximum) || minimum > maximum) {
			throw new Error("XLSX worksheet contains an invalid column definition");
		}
		return { column, minimum, maximum };
	});
	const result: XmlObject[] = [];
	for (const definition of definitions) {
		if (definition.maximum < range.start.column || definition.minimum > range.end.column) {
			result.push(definition.column);
			continue;
		}
		if (definition.minimum < range.start.column) {
			const left = clone(definition.column);
			setAttr(left, "max", range.start.column - 1);
			result.push(left);
		}
		if (definition.maximum > range.end.column) {
			const right = clone(definition.column);
			setAttr(right, "min", range.end.column + 1);
			result.push(right);
		}
	}
	for (let number = range.start.column; number <= range.end.column; number += 1) {
		const inherited = definitions
			.filter((item) => item.minimum <= number && item.maximum >= number)
			.at(-1);
		const column = inherited ? clone(inherited.column) : {};
		setAttr(column, "min", number);
		setAttr(column, "max", number);
		setAttr(column, "width", width);
		setAttr(column, "customWidth", "1");
		result.push(column);
	}
	result.sort((left, right) => Number(attr(left, "min")) - Number(attr(right, "min")));
	setDirectObjects(columns, "col", result);
	return range.end.column - range.start.column + 1;
}

function hasStylePatch(patch: XlsxPatchSet): boolean {
	return Boolean(
		patch.font ||
			patch.fill !== undefined ||
			patch.alignment ||
			patch.number_format !== undefined ||
			patch.border,
	);
}

function validatePatch(patch: XlsxPatchSet): void {
	if (!hasStylePatch(patch) && patch.column_width === undefined && patch.row_height === undefined) {
		throw new Error("XLSX patch requires at least one change");
	}
}

function styleSatisfies(styles: XmlObject, styleIndex: number, patch: XlsxPatchSet): boolean {
	const xfs = styleCollection(styles, "cellXfs", "xf");
	const xf = xfs[styleIndex];
	if (!xf) return false;
	if (patch.font) {
		const font = styleCollection(styles, "fonts", "font")[Number(attr(xf, "fontId") ?? "0")];
		if (patch.font.bold !== undefined && (directKey(font, "b") !== undefined) !== patch.font.bold)
			return false;
		if (
			patch.font.italic !== undefined &&
			(directKey(font, "i") !== undefined) !== patch.font.italic
		)
			return false;
		if (
			patch.font.size !== undefined &&
			Number(attr(directObject(font, "sz"), "val")) !== patch.font.size
		)
			return false;
	}
	if (patch.fill !== undefined) {
		const fill = styleCollection(styles, "fills", "fill")[Number(attr(xf, "fillId") ?? "0")];
		const color = directObject(directObject(fill, "patternFill"), "fgColor");
		if (attr(color, "rgb")?.slice(-6).toUpperCase() !== patch.fill.toUpperCase()) return false;
	}
	const alignment = directObject(xf, "alignment");
	if (
		patch.alignment?.horizontal !== undefined &&
		attr(alignment, "horizontal") !== patch.alignment.horizontal
	)
		return false;
	if (
		patch.alignment?.vertical !== undefined &&
		attr(alignment, "vertical") !== patch.alignment.vertical
	)
		return false;
	if (
		patch.alignment?.wrap_text !== undefined &&
		(attr(alignment, "wrapText") === "1") !== patch.alignment.wrap_text
	)
		return false;
	if (patch.number_format !== undefined) {
		const formats = directObjects(directObject(styles, "numFmts"), "numFmt");
		const id = Number(attr(xf, "numFmtId") ?? "0");
		if (
			id >= 164 &&
			attr(
				formats.find((format) => Number(attr(format, "numFmtId")) === id),
				"formatCode",
			) !== patch.number_format
		)
			return false;
	}
	if (patch.border) {
		const border = styleCollection(styles, "borders", "border")[
			Number(attr(xf, "borderId") ?? "0")
		];
		for (const edge of ["left", "right", "top", "bottom"]) {
			if (attr(directObject(border, edge), "style") !== patch.border.style) return false;
		}
	}
	return true;
}

function validatePatchedSheet(
	pkg: Package,
	sheet: ResolvedSheet,
	range: CellRange,
	patch: XlsxPatchSet,
): void {
	const worksheet = rootObject(parseXml(pkg[sheet.part], sheet.part), "worksheet", sheet.part);
	const rows = rowsByNumber(
		worksheet,
		xmlPrefix(parseXml(pkg[sheet.part], sheet.part), "worksheet"),
	);
	let cellXfs = 0;
	let styles: XmlObject | undefined;
	if (hasStylePatch(patch)) {
		styles = rootObject(
			parseXml(pkg["xl/styles.xml"], "xl/styles.xml"),
			"styleSheet",
			"xl/styles.xml",
		);
		cellXfs = styleCollection(styles, "cellXfs", "xf").length;
	}
	if (patch.row_height !== undefined || hasStylePatch(patch)) {
		for (let rowNumber = range.start.row; rowNumber <= range.end.row; rowNumber += 1) {
			const row = rows.get(rowNumber);
			if (!row) throw new Error("XLSX patch validation failed for target rows");
			if (patch.row_height !== undefined && Number(attr(row, "ht")) !== patch.row_height) {
				throw new Error("XLSX patch validation failed for row height");
			}
			if (!hasStylePatch(patch)) continue;
			const cells = cellsByAddress(row);
			for (let column = range.start.column; column <= range.end.column; column += 1) {
				const cell = cells.get(formatCell({ row: rowNumber, column }));
				const styleIndex = Number(attr(cell, "s"));
				if (
					!cell ||
					!Number.isInteger(styleIndex) ||
					styleIndex < 0 ||
					styleIndex >= cellXfs ||
					!styles ||
					!styleSatisfies(styles, styleIndex, patch)
				) {
					throw new Error("XLSX patch validation failed for cell styles");
				}
			}
		}
	}
	if (patch.column_width !== undefined) {
		const definitions = directObjects(directObject(worksheet, "cols"), "col");
		for (let column = range.start.column; column <= range.end.column; column += 1) {
			const definition = definitions
				.filter(
					(candidate) =>
						Number(attr(candidate, "min")) <= column && Number(attr(candidate, "max")) >= column,
				)
				.at(-1);
			if (!definition || Number(attr(definition, "width")) !== patch.column_width) {
				throw new Error("XLSX patch validation failed for column width");
			}
		}
	}
}

export async function patchXlsx(
	inputPath: string,
	outputPath: string,
	patch: XlsxPatch,
	dryRun = false,
	signal?: AbortSignal,
): Promise<XlsxPatchResult> {
	const input = resolve(inputPath);
	const output = resolve(outputPath);
	if (input === output) throw new Error("XLSX patch output must differ from input");
	validatePatch(patch.patches);
	const range = parseRange(patch.range);
	if (
		(hasStylePatch(patch.patches) && cellCount(range) > MAX_MUTATED_CELLS) ||
		(patch.patches.row_height !== undefined &&
			range.end.row - range.start.row + 1 > MAX_MUTATED_CELLS)
	) {
		throw new Error("XLSX range exceeds the supported mutation limit");
	}
	throwIfAborted(signal);
	const pkg = await load(input);
	validatePackage(pkg);
	const sheet = resolveSheet(pkg, patch.sheet);
	const changedParts = new Set<string>([sheet.part]);
	const result: XlsxPatchResult = {
		sheet: sheet.name,
		range: `${formatCell(range.start)}:${formatCell(range.end)}`,
		styledCells: hasStylePatch(patch.patches) ? cellCount(range) : 0,
		columnsChanged:
			patch.patches.column_width === undefined ? 0 : range.end.column - range.start.column + 1,
		rowsChanged: patch.patches.row_height === undefined ? 0 : range.end.row - range.start.row + 1,
		changedParts: [],
		output,
		wouldWrite: true,
		written: false,
	};
	if (hasStylePatch(patch.patches)) changedParts.add("xl/styles.xml");
	result.changedParts = [...changedParts].sort();
	const worksheetDocument = parseXml(pkg[sheet.part], sheet.part);
	const worksheet = rootObject(worksheetDocument, "worksheet", sheet.part);
	const rows = rowsByNumber(worksheet, xmlPrefix(worksheetDocument, "worksheet"));
	let stylesDocument: XmlObject | undefined;
	let styles: XmlObject | undefined;
	const styleIndexes = new Map<number, number>();
	if (hasStylePatch(patch.patches)) {
		if (!pkg["xl/styles.xml"]) throw new Error("XLSX style patch requires xl/styles.xml");
		stylesDocument = parseXml(pkg["xl/styles.xml"], "xl/styles.xml");
		styles = rootObject(stylesDocument, "styleSheet", "xl/styles.xml");
	}
	if (patch.patches.row_height !== undefined || styles) {
		for (let rowNumber = range.start.row; rowNumber <= range.end.row; rowNumber += 1) {
			throwIfAborted(signal);
			const row = ensureRow(worksheet, rows, rowNumber);
			if (patch.patches.row_height !== undefined) {
				setAttr(row, "ht", patch.patches.row_height);
				setAttr(row, "customHeight", "1");
			}
			if (!styles) continue;
			const cells = cellsByAddress(row);
			for (let column = range.start.column; column <= range.end.column; column += 1) {
				const cell = ensureCell(cells, formatCell({ row: rowNumber, column }));
				const baseIndex = effectiveStyleIndex(worksheet, row, cell, column);
				let styleIndex = styleIndexes.get(baseIndex);
				if (styleIndex === undefined) {
					styleIndex = styleIndexFor(styles, baseIndex, patch.patches);
					styleIndexes.set(baseIndex, styleIndex);
				}
				setAttr(cell, "s", styleIndex);
			}
			syncCells(row, cells, xmlPrefix(worksheetDocument, "worksheet"));
		}
		syncRows(worksheet, rows, xmlPrefix(worksheetDocument, "worksheet"));
	}
	if (patch.patches.column_width !== undefined) {
		patchColumns(worksheet, range, patch.patches.column_width);
	}
	updateDimension(worksheet, xmlPrefix(worksheetDocument, "worksheet"));
	pkg[sheet.part] = serializeXml(worksheetDocument);
	if (stylesDocument) pkg["xl/styles.xml"] = serializeXml(stylesDocument);
	validatePackage(pkg);
	validatePatchedSheet(pkg, sheet, range, patch.patches);
	throwIfAborted(signal);
	if (dryRun) return result;
	await writePackageAtomically(
		pkg,
		output,
		async (temporary) => {
			const candidate = await load(temporary);
			validatePackage(candidate);
			const candidateSheet = resolveSheet(candidate, sheet.name);
			validatePatchedSheet(candidate, candidateSheet, range, patch.patches);
		},
		signal,
	);
	result.written = true;
	return result;
}
