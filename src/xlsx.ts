import { randomUUID } from "node:crypto";
import {
	copyFile,
	mkdir,
	mkdtemp,
	readFile,
	realpath,
	rename,
	rm,
	writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, join, posix, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { parse as parseCsv } from "csv-parse/sync";
import { XMLBuilder, XMLParser } from "fast-xml-parser";
import { type OfficePackage, packOfficePackage, unpackOfficePackage } from "./ooxml.js";
import { commandError, runCommand } from "./process.js";

const MAX_COLUMN = 16_384;
const MAX_ROW = 1_048_576;
const MAX_MUTATED_CELLS = 1_000_000;
const MAX_STRUCTURE_RANGES = 1_000;

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
	mergedRangeAddresses: string[];
	mergedRangesTruncated: boolean;
	hiddenRowRanges: string[];
	hiddenRowsTruncated: boolean;
	hiddenColumnRanges: string[];
	hiddenColumnsTruncated: boolean;
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
	columnMap?: Record<string, string>;
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
		const xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
		if (/[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u.test(xml)) {
			throw new Error("Invalid XML characters");
		}
		return parser.parse(xml, true) as XmlObject;
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

function formatRange(range: CellRange): string {
	const start = formatCell(range.start);
	const end = formatCell(range.end);
	return start === end ? start : `${start}:${end}`;
}

function mergedRanges(worksheet: XmlObject): CellRange[] {
	return directObjects(directObject(worksheet, "mergeCells"), "mergeCell").map((merge) => {
		const reference = attr(merge, "ref");
		if (!reference) throw new Error("XLSX worksheet contains a merged range without a reference");
		try {
			return parseRange(reference);
		} catch {
			throw new Error("XLSX worksheet contains an invalid merged range");
		}
	});
}

function boundedRanges(ranges: string[]): { values: string[]; truncated: boolean } {
	return {
		values: ranges.slice(0, MAX_STRUCTURE_RANGES),
		truncated: ranges.length > MAX_STRUCTURE_RANGES,
	};
}

function hiddenRowRanges(worksheet: XmlObject): string[] {
	const rows = directObjects(sheetData(worksheet), "row")
		.filter((row) => ["1", "true"].includes(attr(row, "hidden") ?? ""))
		.map((row) => Number(attr(row, "r")))
		.filter((row) => Number.isInteger(row) && row >= 1 && row <= MAX_ROW)
		.sort((left, right) => left - right);
	const ranges: string[] = [];
	for (let index = 0; index < rows.length; ) {
		const start = rows[index];
		let end = start;
		while (rows[index + 1] === end + 1) end = rows[++index];
		ranges.push(start === end ? String(start) : `${start}:${end}`);
		index += 1;
	}
	return ranges;
}

function hiddenColumnRanges(worksheet: XmlObject): string[] {
	return directObjects(directObject(worksheet, "cols"), "col")
		.filter((column) => ["1", "true"].includes(attr(column, "hidden") ?? ""))
		.map((column) => {
			const start = Number(attr(column, "min"));
			const end = Number(attr(column, "max"));
			if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end > MAX_COLUMN) {
				throw new Error("XLSX worksheet contains an invalid hidden column definition");
			}
			return start === end
				? columnNumberToName(start)
				: `${columnNumberToName(start)}:${columnNumberToName(end)}`;
		});
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
	for (const range of mergedRanges(worksheet)) {
		minimumRow = Math.min(minimumRow, range.start.row);
		minimumColumn = Math.min(minimumColumn, range.start.column);
		maximumRow = Math.max(maximumRow, range.end.row);
		maximumColumn = Math.max(maximumColumn, range.end.column);
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
	const merges = boundedRanges(mergedRanges(worksheet).map(formatRange));
	const hiddenRows = boundedRanges(hiddenRowRanges(worksheet));
	const hiddenColumns = boundedRanges(hiddenColumnRanges(worksheet));
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
		mergedRanges: mergedRanges(worksheet).length,
		mergedRangeAddresses: merges.values,
		mergedRangesTruncated: merges.truncated,
		hiddenRowRanges: hiddenRows.values,
		hiddenRowsTruncated: hiddenRows.truncated,
		hiddenColumnRanges: hiddenColumns.values,
		hiddenColumnsTruncated: hiddenColumns.truncated,
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

interface MappedColumn {
	source: number;
	destination: number;
}

function validateFilledPackage(
	pkg: Package,
	sheetName: string,
	destination: CellRange,
	rows: string[][],
	columns: MappedColumn[],
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
		for (const column of columns) {
			const cell = cells.get(formatCell({ row, column: column.destination }));
			const expected = rows[row - destination.start.row][column.source] ?? "";
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

function resolveMappedColumns(
	columnMap: Record<string, string> | undefined,
	header: string[] | undefined,
	sourceColumns: number,
	startColumn: number,
): MappedColumn[] {
	if (!columnMap) {
		return Array.from({ length: sourceColumns }, (_, source) => ({
			source,
			destination: startColumn + source,
		}));
	}
	const entries = Object.entries(columnMap);
	if (entries.length === 0) throw new Error("XLSX column_map must not be empty");
	const ordinal = entries.every(([source]) => /^[1-9]\d*$/.test(source));
	const headers = entries.every(([source]) => !/^[1-9]\d*$/.test(source));
	if (!ordinal && !headers)
		throw new Error("XLSX column_map keys must be all ordinals or all headers");
	const destinations = entries.map(([, destination]) => columnNameToNumber(destination));
	if (new Set(destinations).size !== destinations.length) {
		throw new Error("XLSX column_map destinations must be unique");
	}
	if (ordinal) {
		return entries.map(([source], index) => {
			const sourceIndex = Number(source) - 1;
			if (sourceIndex >= sourceColumns)
				throw new Error("XLSX column_map source ordinal is outside the CSV");
			return { source: sourceIndex, destination: destinations[index] };
		});
	}
	if (!header) throw new Error("Header XLSX column_map requires has_header=true");
	const headerIndexes = new Map<string, number>();
	for (const [index, value] of header.entries()) {
		if (headerIndexes.has(value)) throw new Error("CSV headers must be unique for XLSX column_map");
		headerIndexes.set(value, index);
	}
	return entries.map(([source], index) => {
		const sourceIndex = headerIndexes.get(source);
		if (sourceIndex === undefined)
			throw new Error("XLSX column_map header is not present in the CSV");
		return { source: sourceIndex, destination: destinations[index] };
	});
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
	const hasHeader = options.hasHeader ?? true;
	const rows = hasHeader ? allRows.slice(1) : allRows;
	if (rows.length === 0) throw new Error("CSV contains no destination data rows");
	const sourceColumns = Math.max(0, ...allRows.map((row) => row.length));
	if (sourceColumns === 0) throw new Error("CSV contains no destination columns");
	const mappedColumns = resolveMappedColumns(
		options.columnMap,
		hasHeader ? allRows[0] : undefined,
		sourceColumns,
		start.column,
	);
	const minimumColumn = Math.min(...mappedColumns.map((column) => column.destination));
	const maximumColumn = Math.max(...mappedColumns.map((column) => column.destination));
	const destination = {
		start: { row: start.row, column: minimumColumn },
		end: { row: start.row + rows.length - 1, column: maximumColumn },
	};
	if (rows.length * mappedColumns.length > MAX_MUTATED_CELLS) {
		throw new Error("XLSX destination exceeds the supported mutation limit");
	}
	const targetRange = formatRange(destination);
	const worksheetDocument = parseXml(pkg[sheet.part], sheet.part);
	const worksheet = rootObject(worksheetDocument, "worksheet", sheet.part);
	const prefix = xmlPrefix(worksheetDocument, "worksheet");
	const existingRows = rowsByNumber(worksheet, prefix);
	const mappedDestinations = mappedColumns.flatMap((column) =>
		rows.map((_, rowOffset) => ({
			start: { row: start.row + rowOffset, column: column.destination },
			end: { row: start.row + rowOffset, column: column.destination },
		})),
	);
	let formulaConflicts = 0;
	for (const formulaRange of formulaRanges(worksheet)) {
		if (mappedDestinations.some((cell) => overlaps(cell, formulaRange))) formulaConflicts += 1;
	}
	if (formulaConflicts > 0) {
		throw new Error(
			`Destination intersects ${formulaConflicts} formula cell${formulaConflicts === 1 ? "" : "s"}. Refusing to overwrite formulas.`,
		);
	}
	let mergedConflicts = 0;
	for (const mergedRange of mergedRanges(worksheet)) {
		if (mappedDestinations.some((cell) => overlaps(cell, mergedRange))) mergedConflicts += 1;
	}
	if (mergedConflicts > 0) {
		throw new Error(
			`Destination intersects ${mergedConflicts} merged range${mergedConflicts === 1 ? "" : "s"}. Refusing to overwrite merged cells.`,
		);
	}
	const result: XlsxWriteResult = {
		sheet: sheet.name,
		targetRange,
		writtenRange: targetRange,
		rows: rows.length,
		columns: mappedColumns.length,
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
		for (const column of mappedColumns) {
			const address = formatCell({
				row: start.row + rowOffset,
				column: column.destination,
			});
			setCellValue(ensureCell(cells, address), rows[rowOffset][column.source] ?? "", mode, prefix);
		}
		syncCells(row, cells, prefix);
	}
	syncRows(worksheet, existingRows, prefix);
	updateDimension(worksheet, prefix);
	pkg[sheet.part] = serializeXml(worksheetDocument);
	validatePackage(pkg);
	validateFilledPackage(pkg, sheet.name, destination, rows, mappedColumns, mode);
	throwIfAborted(options.signal);
	if (options.dryRun) return result;
	await writePackageAtomically(
		pkg,
		output,
		async (temporary) => {
			const candidate = await load(temporary);
			validatePackage(candidate);
			validateFilledPackage(candidate, sheet.name, destination, rows, mappedColumns, mode);
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

export interface XlsxFormula {
	cell: string;
	formula: string;
}

export function validateXlsxFormula(value: unknown): asserts value is XlsxFormula {
	if (
		!isObject(value) ||
		Object.keys(value).some((key) => key !== "cell" && key !== "formula") ||
		typeof value.cell !== "string" ||
		typeof value.formula !== "string" ||
		!value.formula.startsWith("=") ||
		!value.formula.slice(1).trim() ||
		value.formula.includes("\r") ||
		value.formula.length > 8192 ||
		/[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u.test(value.formula)
	) {
		throw new Error(
			"Invalid XLSX formula request: use a single cell and an XML-safe formula starting with = (maximum 8192 characters)",
		);
	}
	parseCellAddress(value.cell);
}

export function validateXlsxFormulas(value: unknown): asserts value is XlsxFormula[] {
	if (!Array.isArray(value) || value.length === 0 || value.length > 1000) {
		throw new Error("XLSX formula batch must contain 1 to 1000 targets");
	}
	const seen = new Set<string>();
	for (const request of value) {
		validateXlsxFormula(request);
		const cell = formatCell(parseCellAddress(request.cell));
		if (seen.has(cell)) throw new Error("XLSX formula batch contains duplicate cells");
		seen.add(cell);
	}
}

export interface XlsxFormulaSheet {
	sheet: string;
	formulas: XlsxFormula[];
}

export function validateXlsxFormulaSheets(value: unknown): asserts value is XlsxFormulaSheet[] {
	if (!Array.isArray(value) || value.length === 0 || value.length > 1000)
		throw new Error("XLSX formula sheets must contain 1 to 1000 worksheet targets");
	const seen = new Set<string>();
	let count = 0;
	for (const group of value) {
		if (
			!isObject(group) ||
			Object.keys(group).some((key) => key !== "sheet" && key !== "formulas") ||
			typeof group.sheet !== "string" ||
			!group.sheet.trim()
		)
			throw new Error("Invalid XLSX formula worksheet request");
		if (seen.has(group.sheet)) throw new Error("Duplicate XLSX formula worksheet target");
		seen.add(group.sheet);
		validateXlsxFormulas(group.formulas);
		count += group.formulas.length;
		if (count > 1000) throw new Error("XLSX formula batch exceeds 1000 total targets");
	}
}

// This exact Calc extension contains only an empty element and literal attributes.
// Keep unknown extensions rejected: they may carry formula or mixed-content metadata.
function isLibreOfficeCalcExtension(value: unknown): boolean {
	return (
		stable(value) ===
		stable({
			ext: {
				"@_xmlns:loext": "http://schemas.libreoffice.org/",
				"@_uri": "{7626C862-2A13-11E5-B345-FEFF819CDC9F}",
				"loext:extCalcPr": { "@_stringRefSyntax": "CalcA1ExcelA1" },
			},
		})
	);
}

// Only simple, uniformly named SpreadsheetML (plus the exact Calc extension)
// can safely use the object builder. It loses mixed order and numeric entities.
function assertFormulaEditableXml(bytes: Uint8Array): void {
	let xml: string;
	try {
		xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
	} catch {
		throw new Error("Unsupported XLSX formula XML text representation");
	}
	const withoutDeclaration = xml.replace(/^<\?xml(?=\s)[\s\S]*?\?>/, "");
	if (
		/[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u.test(xml) ||
		/&#|\r|<!\[CDATA\[|<!--|<!DOCTYPE/i.test(xml) ||
		withoutDeclaration.includes("<?")
	) {
		throw new Error("Unsupported XLSX formula XML text representation");
	}
	const document = parseXml(bytes, "");
	const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const rootKey = Object.keys(document).find((key) => !key.startsWith("?"));
	if (!rootKey) throw new Error("Ambiguous XLSX XML structure");
	const prefix = rootKey.includes(":") ? rootKey.split(":")[0] : "";
	const visit = (value: unknown, inherited: Record<string, string>): void => {
		if (Array.isArray(value)) {
			for (const child of value) visit(child, inherited);
			return;
		}
		if (!isObject(value)) return;
		const namespaces = { ...inherited };
		for (const [key, child] of Object.entries(value)) {
			if (key === "@_xmlns") namespaces[""] = String(child);
			else if (key.startsWith("@_xmlns:")) namespaces[key.slice(8)] = String(child);
		}
		// Namespace rebinding and foreign defaults can misnamespace newly inserted nodes.
		if (namespaces[prefix] !== main || (namespaces[""] && namespaces[""] !== main)) {
			throw new Error("Unsupported XLSX formula XML namespace");
		}
		for (const key of Object.keys(value)) {
			if (!key.startsWith("@_") || key.startsWith("@_xmlns")) continue;
			const name = key.slice(2);
			if (!name.includes(":")) continue;
			const attributePrefix = name.split(":")[0];
			if (
				!(attributePrefix === "xml" && name === "xml:space") &&
				!(
					name.endsWith(":id") &&
					namespaces[attributePrefix] ===
						"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
				)
			) {
				throw new Error("Unsupported XLSX formula XML attribute namespace");
			}
		}
		const children = Object.keys(value).filter((key) => !key.startsWith("@_") && key !== "#text");
		if (children.length && String(value["#text"] ?? "").trim()) {
			throw new Error("Unsupported XLSX formula mixed XML content");
		}
		for (const key of children) {
			const childPrefix = key.includes(":") ? key.split(":")[0] : "";
			if (
				prefix === "" &&
				childPrefix === "" &&
				localName(key) === "extLst" &&
				isLibreOfficeCalcExtension(value[key])
			) {
				continue;
			}
			if (childPrefix !== prefix || localName(key) === "AlternateContent") {
				throw new Error("Unsupported XLSX formula XML namespace");
			}
			visit(value[key], namespaces);
		}
	};
	visit(document[rootKey], {});
}

// Resolve existing ancestors without creating directories, including during dry runs.
async function canonicalFormulaOutput(input: string, output: string): Promise<string> {
	const canonical = async (path: string): Promise<string> => {
		try {
			return await realpath(path);
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
			return join(await canonical(dirname(path)), basename(path));
		}
	};
	const source = await realpath(input);
	const destination = join(await canonical(dirname(output)), basename(output));
	if (source === destination || source === (await canonical(output))) {
		throw new Error("XLSX patch output must differ from input");
	}
	return destination;
}

// The ordinary object parser groups repeated tags. Reject ambiguous singleton
// elements and alternate-prefix collisions before using the existing mutation helpers.
function assertFormulaTree(value: unknown): void {
	if (Array.isArray(value)) {
		for (const child of value) assertFormulaTree(child);
		return;
	}
	if (!isObject(value)) return;
	const names = new Set<string>();
	for (const [key, child] of Object.entries(value)) {
		const name = localName(key);
		if (names.has(name)) throw new Error("Ambiguous XLSX XML structure");
		names.add(name);
		if (
			(name === "extLst" && !isLibreOfficeCalcExtension(child)) ||
			["metadata", "dynamicArrayProperties"].includes(name)
		) {
			throw new Error("Unsupported XLSX formula metadata");
		}
		if (
			["sheetData", "dimension", "calcPr", "sheets", "f", "v", "is"].includes(name) &&
			Array.isArray(child)
		) {
			throw new Error("Ambiguous XLSX XML structure");
		}
		assertFormulaTree(child);
	}
}

function validateFormulaWorksheet(worksheet: XmlObject): void {
	assertFormulaTree(worksheet);
	const dataKey = directKey(worksheet, "sheetData");
	if (!dataKey) throw new Error("XLSX formula editing requires explicit sheetData");
	if (!isObject(worksheet[dataKey]) && worksheet[dataKey] !== "") {
		throw new Error("Ambiguous XLSX row structure");
	}
	const data = sheetData(worksheet);
	if (
		Object.keys(data).some(
			(key) => !key.startsWith("@_") && !["row", "#text"].includes(localName(key)),
		)
	) {
		throw new Error("Ambiguous XLSX row structure");
	}
	const rowKey = directKey(data, "row");
	const rawRows = rowKey ? (Array.isArray(data[rowKey]) ? data[rowKey] : [data[rowKey]]) : [];
	if ((rawRows as unknown[]).some((row) => !isObject(row)))
		throw new Error("Ambiguous XLSX row structure");
	let previousRow = 0;
	for (const row of directObjects(data, "row")) {
		const reference = attr(row, "r") ?? "";
		const number = Number(reference);
		if (!/^[1-9]\d*$/.test(reference) || number <= previousRow || number > MAX_ROW) {
			throw new Error("Ambiguous XLSX row structure");
		}
		previousRow = number;
		if (
			Object.keys(row).some(
				(key) => !key.startsWith("@_") && !["c", "#text"].includes(localName(key)),
			)
		) {
			throw new Error("Ambiguous XLSX cell structure");
		}
		const cellKey = directKey(row, "c");
		const rawCells = cellKey ? (Array.isArray(row[cellKey]) ? row[cellKey] : [row[cellKey]]) : [];
		if ((rawCells as unknown[]).some((cell) => !isObject(cell)))
			throw new Error("Ambiguous XLSX cell structure");
		let previousColumn = 0;
		for (const cell of directObjects(row, "c")) {
			const address = parseCellAddress(attr(cell, "r") ?? "");
			if (address.row !== number || address.column <= previousColumn) {
				throw new Error("Ambiguous XLSX cell structure");
			}
			previousColumn = address.column;
			if (attr(cell, "cm") !== undefined || attr(cell, "vm") !== undefined) {
				throw new Error("Unsupported XLSX formula metadata");
			}
			const key = directKey(cell, "f");
			if (!key) continue;
			const formula = cell[key];
			if (isObject(formula)) {
				for (const [name, setting] of Object.entries(formula)) {
					// aca=false is the ordinary-formula default, emitted explicitly by Calc.
					if (
						name !== "#text" &&
						!(name === "@_t" && setting === "normal") &&
						!(name === "@_aca" && ["false", "0"].includes(String(setting)))
					) {
						throw new Error("XLSX formula editing supports only ordinary formulas");
					}
				}
			} else if (typeof formula !== "string") {
				throw new Error("XLSX formula editing supports only ordinary formulas");
			}
		}
	}
}

// The object helpers select one spelling per local name. Require a uniform
// element prefix and resolve every flat metadata child's namespace before use.
function assertFormulaMetadata(
	document: XmlObject,
	name: string,
	namespace: string,
	children: string[],
): void {
	const fail = (): never => {
		throw new Error("Unsupported XLSX package metadata namespace structure");
	};
	const roots = Object.keys(document).filter((key) => key !== "?xml");
	if (roots.length !== 1 || localName(roots[0]) !== name) fail();
	const root = document[roots[0]];
	if (!isObject(root)) {
		throw new Error("Unsupported XLSX package metadata namespace structure");
	}
	const prefix = roots[0].includes(":") ? roots[0].split(":")[0] : "";
	const declaration = prefix ? `@_xmlns:${prefix}` : "@_xmlns";
	if (root[declaration] !== namespace) fail();
	for (const [key, value] of Object.entries(root)) {
		if (key.startsWith("@_") || (key === "#text" && !String(value).trim())) continue;
		const childPrefix = key.includes(":") ? key.split(":")[0] : "";
		if (childPrefix !== prefix || !children.includes(localName(key))) fail();
		for (const child of Array.isArray(value) ? value : [value]) {
			if (!isObject(child) || (child[declaration] ?? root[declaration]) !== namespace) fail();
		}
	}
}

function validateFormulaPackage(pkg: Package, allowCalculationChain = false): void {
	validatePackage(pkg);
	const sheets = workbookSheets(pkg).sheets;
	if (
		new Set(sheets.map((sheet) => sheet.name.toLowerCase())).size !== sheets.length ||
		new Set(sheets.map((sheet) => sheet.part)).size !== sheets.length
	) {
		throw new Error("Ambiguous XLSX sheet structure");
	}
	for (const [part, bytes] of Object.entries(pkg)) {
		if (
			/^_xmlsignatures\//i.test(part) ||
			(!allowCalculationChain && /(?:^|\/)calcChain\.xml$/i.test(part)) ||
			/(?:^|\/)metadata\.xml$/i.test(part)
		) {
			throw new Error("Unsupported XLSX signatures, calculation chain, or formula metadata");
		}
		if (!part.endsWith(".xml") && !part.endsWith(".rels")) continue;
		const document = parseXml(bytes, part);
		if (part.endsWith(".rels")) {
			assertFormulaMetadata(
				document,
				"Relationships",
				"http://schemas.openxmlformats.org/package/2006/relationships",
				["Relationship"],
			);
			const ids = new Set<string>();
			for (const relationship of directObjects(
				rootObject(document, "Relationships", part),
				"Relationship",
			)) {
				const id = attr(relationship, "Id");
				if (!id || ids.has(id)) throw new Error("Ambiguous XLSX relationship structure");
				ids.add(id);
				if (
					/\/sheetMetadata$|digital-signature/i.test(attr(relationship, "Type") ?? "") ||
					(!allowCalculationChain && /\/calcChain$/i.test(attr(relationship, "Type") ?? ""))
				) {
					throw new Error("Unsupported XLSX signatures, calculation chain, or formula metadata");
				}
			}
		}
		if (part === "[Content_Types].xml") {
			assertFormulaMetadata(
				document,
				"Types",
				"http://schemas.openxmlformats.org/package/2006/content-types",
				["Default", "Override"],
			);
			const types = rootObject(document, "Types", part);
			for (const override of [
				...directObjects(types, "Override"),
				...directObjects(types, "Default"),
			]) {
				if (
					/digital-signature|sheetMetadata/i.test(attr(override, "ContentType") ?? "") ||
					(!allowCalculationChain && /calcChain/i.test(attr(override, "ContentType") ?? ""))
				) {
					throw new Error("Unsupported XLSX signatures, calculation chain, or formula metadata");
				}
			}
		}
		if (directKey(document, "worksheet")) {
			validateFormulaWorksheet(rootObject(document, "worksheet", part));
		}
	}
}

// A calculation chain is disposable calculation order, not formula content.
// Support only the conventional, uniquely linked part; never guess at aliases.
function discardCalculationChain(pkg: Package): string[] {
	const chain = "xl/calcChain.xml";
	const relsPart = "xl/_rels/workbook.xml.rels";
	const typesPart = "[Content_Types].xml";
	const chainType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
	const contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
	const fail = (): never => {
		throw new Error("Unsupported XLSX calculation chain structure");
	};
	const incoming: { part: string; relationship: XmlObject }[] = [];
	for (const part of Object.keys(pkg).filter((name) => name.endsWith(".rels"))) {
		const root = rootObject(parseXml(pkg[part], part), "Relationships", part);
		for (const relationship of directObjects(root, "Relationship")) {
			const type = attr(relationship, "Type") ?? "";
			const target = relationshipTarget(part, attr(relationship, "Target") ?? "");
			if (
				/\/calcChain$/i.test(type) ||
				(attr(relationship, "TargetMode") !== "External" && target === chain)
			)
				incoming.push({ part, relationship });
		}
	}
	const typesDocument = parseXml(pkg[typesPart], typesPart);
	const types = rootObject(typesDocument, "Types", typesPart);
	const overrides = directObjects(types, "Override").filter(
		(item) =>
			/calcChain/i.test(attr(item, "ContentType") ?? "") || attr(item, "PartName") === `/${chain}`,
	);
	const hasChain = Object.keys(pkg).some((part) => /(?:^|\/)calcChain\.xml$/i.test(part));
	if (!hasChain && !incoming.length && !overrides.length) return [];
	if (
		!pkg[chain] ||
		incoming.length !== 1 ||
		overrides.length !== 1 ||
		incoming[0].part !== relsPart ||
		attr(incoming[0].relationship, "Type") !== chainType ||
		attr(incoming[0].relationship, "TargetMode") !== undefined ||
		relationshipTarget(relsPart, attr(incoming[0].relationship, "Target") ?? "") !== chain ||
		attr(overrides[0], "PartName") !== `/${chain}` ||
		attr(overrides[0], "ContentType") !== contentType ||
		pkg["xl/_rels/calcChain.xml.rels"]
	)
		fail();
	// Reserialize only flat, default-namespace package metadata. Reject mixed
	// content and namespace tricks rather than weakening formula XML guards.
	const metadata = (part: string, name: string, namespace: string, children: string[]) => {
		const text = new TextDecoder("utf-8", { fatal: true }).decode(pkg[part]);
		if (/&#|<!|\r|<\?(?!xml\s)/.test(text)) fail();
		const document = parseXml(pkg[part], part);
		if (Object.keys(document).some((key) => key !== name && key !== "?xml")) fail();
		const root = rootObject(document, name, part);
		if (root["@_xmlns"] !== namespace) fail();
		for (const [key, value] of Object.entries(root)) {
			if (key === "@_xmlns" || (key === "#text" && !String(value).trim())) continue;
			if (!children.includes(key)) fail();
			for (const item of Array.isArray(value) ? value : [value]) {
				if (
					!isObject(item) ||
					Object.entries(item).some(
						([attribute, setting]) =>
							attribute === "@_xmlns" ||
							!/^@_[A-Za-z]+$/.test(attribute) ||
							typeof setting !== "string",
					)
				)
					fail();
			}
		}
		return { document, root };
	};
	const rels = metadata(
		relsPart,
		"Relationships",
		"http://schemas.openxmlformats.org/package/2006/relationships",
		["Relationship"],
	);
	const checkedTypes = metadata(
		typesPart,
		"Types",
		"http://schemas.openxmlformats.org/package/2006/content-types",
		["Default", "Override"],
	);
	setDirectObjects(
		rels.root,
		"Relationship",
		directObjects(rels.root, "Relationship").filter(
			(item) => attr(item, "Id") !== attr(incoming[0].relationship, "Id"),
		),
	);
	setDirectObjects(
		checkedTypes.root,
		"Override",
		directObjects(checkedTypes.root, "Override").filter(
			(item) => attr(item, "PartName") !== `/${chain}`,
		),
	);
	pkg[relsPart] = serializeXml(rels.document);
	pkg[typesPart] = serializeXml(checkedTypes.document);
	delete pkg[chain];
	return [chain, relsPart, typesPart];
}

export async function patchXlsxFormula(
	inputPath: string,
	outputPath: string,
	sheetName: string,
	request: XlsxFormula | XlsxFormula[],
	dryRun = false,
	signal?: AbortSignal,
) {
	const batch = Array.isArray(request);
	const result = await patchXlsxFormulaSheets(
		inputPath,
		outputPath,
		[{ sheet: sheetName, formulas: batch ? request : [request] }],
		dryRun,
		signal,
	);
	const { sheets, formulaCount, ...metadata } = result;
	return {
		...metadata,
		sheet: sheets[0].sheet,
		...(batch ? { cells: sheets[0].cells, formulaCount } : { cell: sheets[0].cells[0] }),
	};
}

export async function patchXlsxFormulaSheets(
	inputPath: string,
	outputPath: string,
	requests: XlsxFormulaSheet[],
	dryRun = false,
	signal?: AbortSignal,
) {
	validateXlsxFormulaSheets(requests);
	const input = resolve(inputPath);
	const output = resolve(outputPath);
	if (input === output) throw new Error("XLSX patch output must differ from input");
	throwIfAborted(signal);
	const canonicalOutput = await canonicalFormulaOutput(input, output);
	const pkg = await load(input);
	validateFormulaPackage(pkg, true);
	const chainChanges = discardCalculationChain(pkg);
	validateFormulaPackage(pkg);
	assertFormulaEditableXml(pkg["xl/workbook.xml"]);
	const seenParts = new Set<string>();
	const availableSheets = workbookSheets(pkg).sheets;
	const plans = requests.map((group) => {
		const sheet = availableSheets.find(
			(candidate) => candidate.name.toLowerCase() === group.sheet.toLowerCase(),
		);
		if (!sheet?.worksheet) throw new Error("Requested XLSX worksheet does not exist");
		if (seenParts.has(sheet.part)) throw new Error("Duplicate resolved XLSX worksheet target");
		seenParts.add(sheet.part);
		assertFormulaEditableXml(pkg[sheet.part]);
		const document = parseXml(pkg[sheet.part], sheet.part);
		const worksheet = rootObject(document, "worksheet", sheet.part);
		const prefix = xmlPrefix(document, "worksheet");
		const targets = group.formulas.map((item) => {
			const address = parseCellAddress(item.cell);
			return { address, cellAddress: formatCell(address), formula: item.formula.slice(1) };
		});
		const merges = mergedRanges(worksheet);
		for (const { address } of targets) {
			if (merges.some((range) => overlaps(range, { start: address, end: address }))) {
				throw new Error("XLSX formula target must not be merged");
			}
		}
		return { sheet, document, worksheet, prefix, targets };
	});
	for (const { sheet, document, worksheet, prefix, targets } of plans) {
		const rows = rowsByNumber(worksheet, prefix);
		for (const { address, cellAddress, formula } of targets) {
			const row = ensureRow(worksheet, rows, address.row);
			const cells = cellsByAddress(row);
			const cell = ensureCell(cells, cellAddress);
			removeDirect(cell, "f", "v", "is");
			removeAttr(cell, "t");
			insertDirectBefore(cell, "f", formula, "extLst", prefix);
			syncCells(row, cells, prefix);
		}
		syncRows(worksheet, rows, prefix);
		// Insert dimension before all schema successors, not just sheetData/cols.
		if (!directKey(worksheet, "dimension")) {
			const before = Object.keys(worksheet).find(
				(key) => !key.startsWith("@_") && !["sheetPr", "#text"].includes(localName(key)),
			);
			insertDirectBefore(
				worksheet,
				"dimension",
				{},
				before ? localName(before) : "sheetData",
				prefix,
			);
		}
		updateDimension(worksheet, prefix);
		pkg[sheet.part] = serializeXml(document);
	}
	const workbookPart = "xl/workbook.xml";
	const workbookDocument = parseXml(pkg[workbookPart], workbookPart);
	const workbook = rootObject(workbookDocument, "workbook", workbookPart);
	assertFormulaTree(workbook);
	const calc = directObject(workbook, "calcPr") ?? {};
	const beforeCalc = JSON.stringify(workbook);
	setAttr(calc, "calcMode", "auto");
	setAttr(calc, "fullCalcOnLoad", "1");
	setAttr(calc, "forceFullCalc", "1");
	const successors = [
		"oleSize",
		"customWorkbookViews",
		"pivotCaches",
		"smartTagPr",
		"smartTagTypes",
		"webPublishing",
		"fileRecoveryPr",
		"webPublishObjects",
		"extLst",
	];
	const before = Object.keys(workbook).find((key) => successors.includes(localName(key)));
	removeDirect(workbook, "calcPr");
	insertDirectBefore(
		workbook,
		"calcPr",
		calc,
		before ? localName(before) : "extLst",
		xmlPrefix(workbookDocument, "workbook"),
	);
	const changedParts = [...chainChanges, ...plans.map(({ sheet }) => sheet.part)];
	if (JSON.stringify(workbook) !== beforeCalc) {
		pkg[workbookPart] = serializeXml(workbookDocument);
		changedParts.push(workbookPart);
	}
	const validate = (candidate: Package) => {
		validateFormulaPackage(candidate);
		for (const { sheet, targets } of plans) {
			const checked = rootObject(
				parseXml(candidate[sheet.part], sheet.part),
				"worksheet",
				sheet.part,
			);
			const checkedRows = rowsByNumber(checked);
			for (const { address, cellAddress, formula } of targets) {
				const target = cellsByAddress(checkedRows.get(address.row) ?? {}).get(cellAddress);
				if (
					!target ||
					target[directKey(target, "f") ?? ""] !== formula ||
					directKey(target, "v") ||
					directKey(target, "is") ||
					attr(target, "t") !== undefined
				) {
					throw new Error("XLSX formula output validation failed");
				}
			}
		}
	};
	validate(pkg);
	throwIfAborted(signal);
	if (!dryRun)
		await writePackageAtomically(
			pkg,
			canonicalOutput,
			async (temporary) => validate(await load(temporary)),
			signal,
		);
	return {
		sheets: plans.map(({ sheet, targets }) => ({
			sheet: sheet.name,
			cells: targets.map((target) => target.cellAddress),
			formulaCount: targets.length,
		})),
		formulaCount: plans.reduce((count, plan) => count + plan.targets.length, 0),
		changedParts: changedParts.sort(),
		output,
		wouldWrite: true,
		written: !dryRun,
		warnings: [
			"Other formula cached values are unchanged; stale cache possible until recalculation. No formula evaluation was performed.",
		],
	};
}

// LibreOffice recalculates on load only when told to. This profile forces it for
// OOXML workbooks, so a plain headless conversion yields computed cached values.
const RECALC_PROFILE = `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
`;
const RECALC_TIMEOUT_MS = 180_000;
const CACHE_TYPES = new Set(["b", "str", "e"]);

async function libreOfficeRecalculatedPackage(
	path: string,
	signal?: AbortSignal,
): Promise<Package> {
	const temporary = await mkdtemp(join(tmpdir(), "pi-filler-recalc-"));
	try {
		const profile = join(temporary, "profile");
		await mkdir(join(profile, "user"), { recursive: true });
		await writeFile(join(profile, "user", "registrymodifications.xcu"), RECALC_PROFILE);
		const input = join(temporary, "input.xlsx");
		await copyFile(path, input);
		const converted = await runCommand(
			"libreoffice",
			[
				`-env:UserInstallation=${pathToFileURL(profile).href}`,
				"--headless",
				"--convert-to",
				"xlsx",
				"--outdir",
				join(temporary, "out"),
				input,
			],
			{ signal, cwd: temporary, timeoutMs: RECALC_TIMEOUT_MS },
		);
		if (converted.code !== 0) throw commandError(converted, "libreoffice");
		const bytes = await readFile(join(temporary, "out", "input.xlsx")).catch(() => {
			throw new Error("LibreOffice did not produce a recalculated workbook");
		});
		return unpack(new Uint8Array(bytes));
	} finally {
		await rm(temporary, { recursive: true, force: true });
	}
}

// Formula cells of one worksheet, keyed by normalized address.
function formulaCellsOf(worksheet: XmlObject, prefix?: string): Map<string, XmlObject> {
	const result = new Map<string, XmlObject>();
	for (const row of rowsByNumber(worksheet, prefix).values()) {
		for (const [address, cell] of cellsByAddress(row)) {
			if (directKey(cell, "f")) result.set(address, cell);
		}
	}
	return result;
}

function cachedText(cell: XmlObject): string {
	const key = directKey(cell, "v");
	const value = key ? cell[key] : undefined;
	const text = Array.isArray(value) ? value[0] : value;
	if (text === undefined) return "";
	if (typeof text !== "string") throw new Error("Unsupported XLSX cached value representation");
	return text;
}

function formulaText(cell: XmlObject): string {
	const key = directKey(cell, "f");
	return JSON.stringify(key ? cell[key] : null);
}

export async function recalculateXlsx(
	inputPath: string,
	outputPath: string,
	dryRun = false,
	signal?: AbortSignal,
) {
	const input = resolve(inputPath);
	const output = resolve(outputPath);
	if (input === output) throw new Error("XLSX patch output must differ from input");
	throwIfAborted(signal);
	const canonicalOutput = await canonicalFormulaOutput(input, output);
	const pkg = await load(input);
	validateFormulaPackage(pkg);
	if (Object.keys(pkg).some((part) => /(?:^|\/)externalLinks\//i.test(part))) {
		throw new Error("XLSX external workbook links are unsupported for recalculation");
	}
	const engine = await libreOfficeRecalculatedPackage(input, signal);
	throwIfAborted(signal);
	const computed = new Map<string, Map<string, XmlObject>>();
	for (const sheet of workbookSheets(engine).sheets) {
		if (!sheet.worksheet) continue;
		const document = parseXml(engine[sheet.part], sheet.part);
		const worksheet = rootObject(document, "worksheet", sheet.part);
		computed.set(
			sheet.name.toLowerCase(),
			formulaCellsOf(worksheet, xmlPrefix(document, "worksheet")),
		);
	}
	const plans: { part: string; name: string; formulas: Map<string, string> }[] = [];
	const sheets: { sheet: string; formulaCount: number }[] = [];
	const changedParts: string[] = [];
	let formulaCount = 0;
	for (const sheet of workbookSheets(pkg).sheets) {
		if (!sheet.worksheet) continue;
		assertFormulaEditableXml(pkg[sheet.part]);
		const document = parseXml(pkg[sheet.part], sheet.part);
		const worksheet = rootObject(document, "worksheet", sheet.part);
		const prefix = xmlPrefix(document, "worksheet");
		const cells = formulaCellsOf(worksheet, prefix);
		if (cells.size === 0) continue;
		const source = computed.get(sheet.name.toLowerCase());
		if (!source) throw new Error("Recalculated workbook is missing a worksheet");
		const formulas = new Map<string, string>();
		for (const [address, cell] of cells) {
			const result = source.get(address);
			if (!result) throw new Error("Recalculated workbook is missing a formula cell");
			// Numeric results carry no t attribute; LibreOffice writes t="n" explicitly.
			const rawType = attr(result, "t");
			const type = rawType === undefined || rawType === "n" ? undefined : rawType;
			if (type !== undefined && !CACHE_TYPES.has(type)) {
				throw new Error("Unsupported recalculated XLSX cell type");
			}
			const text = cachedText(result);
			removeDirect(cell, "v");
			removeAttr(cell, "t");
			if (type !== undefined) setAttr(cell, "t", type);
			insertDirectBefore(cell, "v", text, "extLst", prefix);
			formulas.set(address, formulaText(cell));
		}
		plans.push({ part: sheet.part, name: sheet.name, formulas });
		sheets.push({ sheet: sheet.name, formulaCount: cells.size });
		formulaCount += cells.size;
		pkg[sheet.part] = serializeXml(document);
		changedParts.push(sheet.part);
	}
	const validate = (candidate: Package) => {
		validateFormulaPackage(candidate);
		for (const plan of plans) {
			const document = parseXml(candidate[plan.part], plan.part);
			const checked = rootObject(document, "worksheet", plan.part);
			const cells = formulaCellsOf(checked, xmlPrefix(document, "worksheet"));
			if (cells.size !== plan.formulas.size)
				throw new Error("XLSX recalculation validation failed");
			for (const [address, cell] of cells) {
				const expected = plan.formulas.get(address);
				if (expected === undefined || formulaText(cell) !== expected || !directKey(cell, "v")) {
					throw new Error("XLSX recalculation validation failed");
				}
			}
		}
	};
	validate(pkg);
	throwIfAborted(signal);
	const written = !dryRun && formulaCount > 0;
	if (written) {
		await writePackageAtomically(
			pkg,
			canonicalOutput,
			async (temporary) => validate(await load(temporary)),
			signal,
		);
	}
	return {
		sheets,
		formulaCount,
		changedParts: changedParts.sort(),
		output,
		wouldWrite: formulaCount > 0,
		written,
		warnings: [
			"Cached formula results were calculated by LibreOffice, not Microsoft Excel; function support, locale and volatile functions can differ.",
			"Only cached values of existing formula cells were replaced; formulas and inputs are unchanged. Worksheet XML is re-serialized.",
			"Source was not modified. Computed values are not returned.",
		],
	};
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
