import type { FillerInput } from "../src/index.js";

// Compiled by typecheck: enum precision and optionality must survive upstream updates.
type Equal<A, B> =
	(<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type Assert<T extends true> = T;

export type SchemaTypeChecks = [
	Assert<Equal<FillerInput["format"], "docx" | "pdf" | "pptx" | "xlsx">>,
	Assert<Equal<FillerInput["action"], "read" | "search" | "write" | "patch">>,
	Assert<
		Equal<FillerInput["view"], "text" | "formatting" | "image" | "structure" | "review" | undefined>
	>,
	Assert<Equal<FillerInput["track_changes"], "accept" | "reject" | "all" | undefined>>,
	Assert<Equal<FillerInput["value_mode"], "text" | "auto" | undefined>>,
	Assert<
		Equal<
			NonNullable<NonNullable<FillerInput["patches"]>["page"]>["orientation"],
			"portrait" | "landscape" | undefined
		>
	>,
	Assert<
		Equal<
			NonNullable<NonNullable<FillerInput["patches"]>["lineNumbering"]>["mode"],
			"off" | "continuous" | "newPage" | "newSection" | undefined
		>
	>,
	Assert<
		Equal<
			NonNullable<NonNullable<FillerInput["xlsx_patches"]>["alignment"]>["horizontal"],
			"left" | "center" | "right" | undefined
		>
	>,
	Assert<
		Equal<
			NonNullable<NonNullable<FillerInput["xlsx_patches"]>["alignment"]>["vertical"],
			"top" | "center" | "bottom" | undefined
		>
	>,
	Assert<
		Equal<
			NonNullable<NonNullable<FillerInput["xlsx_patches"]>["border"]>["style"],
			"thin" | "medium" | "thick" | undefined
		>
	>,
	Assert<{ format: "docx"; action: "read"; path: string } extends FillerInput ? true : false>,
];
