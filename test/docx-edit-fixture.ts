import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { unzipSync, zipSync } from "fflate";
import { CT, R, REL, W } from "../src/docx-package.js";

export const enc = (value: string) => new TextEncoder().encode(value);
export const p = (value: string) =>
	`<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;
export const section =
	'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="900" w:bottom="1000" w:left="1100" w:right="1200"/></w:sectPr>';
export const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="48"/></w:rPr></w:style></w:styles>`;
export const row = (label: string) =>
	`<w:tr><w:trPr><w:trHeight w:val="180"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="300"/><w:gridSpan w:val="2"/></w:tcPr>${p(label)}</w:tc><w:tc><w:tcPr><w:shd w:fill="000000"/><w:tcBorders><w:top w:val="single"/></w:tcBorders></w:tcPr>${p("Cell content")}</w:tc></w:tr>`;
export const drawing = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5400000" cy="7200000"/><wp:docPr id="1" name="Plot"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Plot"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="5400000" cy="7200000"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
export async function directory() {
	return mkdtemp(join(tmpdir(), "pi-filler-docx-edit-test-"));
}
export async function fixture(
	path: string,
	body: string,
	extras: Record<string, string | Uint8Array> = {},
	rels = "",
) {
	const pkg: Record<string, Uint8Array> = {
		"[Content_Types].xml": enc(
			`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
		),
		"_rels/.rels": enc(
			`<Relationships xmlns="${REL}"><Relationship Id="rIdDocument" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
		),
		"word/document.xml": enc(
			`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${section}</w:body></w:document>`,
		),
		"word/_rels/document.xml.rels": enc(
			`<Relationships xmlns="${REL}"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>${rels}</Relationships>`,
		),
		"word/styles.xml": enc(styles),
	};
	for (const [name, value] of Object.entries(extras))
		pkg[name] = typeof value === "string" ? enc(value) : value;
	await writeFile(path, zipSync(pkg));
	return pkg;
}
export async function packageAt(path: string) {
	return unzipSync(new Uint8Array(await readFile(path)));
}
export async function documentAt(path: string) {
	return new TextDecoder().decode((await packageAt(path))["word/document.xml"]);
}
