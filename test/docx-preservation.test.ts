import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { XMLParser } from "fast-xml-parser";
import { unzipSync, zipSync } from "fflate";
import { inspectDocx, patchDocx } from "../src/docx.js";

const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const document = `<?xml version="1.0"?>
<x:document xmlns:x="${word}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
<x:body>
<x:p><x:pPr><x:sectPr><x:pgSz x:w="1000" x:h="2000"/><x:pgMar x:top="100"/>
<x:lnNumType x:restart="continuous"/><x:pgNumType x:start="1"/></x:sectPr></x:pPr>
<x:r><x:t xml:space="preserve">Alpha </x:t></x:r><x:bookmarkStart x:id="001" x:name="mark"/>
<x:r><x:rPr><x:b/></x:rPr><x:t>BETA</x:t><x:tab/><x:t xml:space="preserve">  tabbed </x:t><x:br/></x:r>
<x:commentRangeStart x:id="2"/><x:ins x:id="3" x:author="Editor"><x:r><x:t xml:space="preserve"> gamma</x:t></x:r></x:ins>
<x:commentRangeEnd x:id="2"/><x:r><x:commentReference x:id="2"/></x:r><x:bookmarkEnd x:id="001"/>
<m:oMath><m:r><m:t>001.230</m:t></m:r></m:oMath><x:fldSimple x:instr="  REF 001  "/>
<!-- preserve this comment --><x:customXml><x:r><x:t><![CDATA[ <literal> & text ]]></x:t></x:r></x:customXml></x:p>
<x:tbl><x:tr><x:tc><x:p><x:r><x:t>001.230</x:t></x:r></x:p></x:tc></x:tr></x:tbl>
<x:p><x:r><x:t>001.230</x:t></x:r><x:del x:id="4"><x:r><x:delText xml:space="preserve"> gone </x:delText></x:r></x:del></x:p>
<x:sectPr><x:pgMar x:top="999"/></x:sectPr></x:body></x:document>`;
const core = `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title xml:lang="en">  001.230  </dc:title><dc:creator> Alice </dc:creator><cp:keywords>001.230</cp:keywords>
<dc:description>A &amp; B</dc:description></cp:coreProperties>`;
const encode = (text: string) => new TextEncoder().encode(text);
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const parser = new XMLParser({
	preserveOrder: true,
	ignoreAttributes: false,
	parseTagValue: false,
	parseAttributeValue: false,
	trimValues: false,
	commentPropName: "#comment",
	cdataPropName: "#cdata",
});
const normalized = (text: string) => parser.parse(text);

async function fixture(t: TestContext) {
	const dir = await mkdtemp(join(tmpdir(), "docx-preserve-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	const input = join(dir, "input.docx");
	const output = join(dir, "output.docx");
	const files: Record<string, Uint8Array> = {
		"[Content_Types].xml": encode("<Types/>"),
		"word/document.xml": encode(document),
		"docProps/core.xml": encode(core),
		"word/comments.xml": encode(
			`<x:comments xmlns:x="${word}"><x:comment x:id="2" x:author="  001.230  "><x:p><x:r><x:t> comment </x:t></x:r></x:p></x:comment></x:comments>`,
		),
		"custom/data.bin": Uint8Array.from([0, 1, 255]),
	};
	await writeFile(input, zipSync(files));
	return { input, output, files };
}

test("document patches preserve all mixed content and later sections", async (t) => {
	for (const [patch, expected] of [
		[{ margins: { top: 720 } }, document.replace('x:top="100"', 'x:top="720"')],
		[{ page: { width: 1234 } }, document.replace('x:w="1000"', 'x:w="1234"')],
		[
			{ lineNumbering: { mode: "newPage" as const } },
			document.replace('x:restart="continuous"', 'x:restart="newPage"'),
		],
		[{ pageNumbering: { start: 7 } }, document.replace('x:start="1"', 'x:start="7"')],
		[
			{ lineNumbering: { mode: "off" as const } },
			document.replace('<x:lnNumType x:restart="continuous"/>', ""),
		],
	] as const) {
		const { input, output, files } = await fixture(t);
		const source = await readFile(input);
		const result = await patchDocx(input, output, patch);
		assert.deepEqual(result.changedParts, ["word/document.xml"]);
		const parts = unzipSync(await readFile(output));
		assert.deepEqual(normalized(decode(parts["word/document.xml"])), normalized(expected));
		for (const [name, bytes] of Object.entries(files)) {
			if (name !== "word/document.xml") assert.deepEqual(parts[name], bytes);
		}
		assert.deepEqual(await readFile(input), source);
	}
});

test("metadata patches preserve literal values and every other part", async (t) => {
	const { input, output, files } = await fixture(t);
	await patchDocx(input, output, { metadata: { creator: "  Bob  " } });
	const parts = unzipSync(await readFile(output));
	assert.deepEqual(
		normalized(decode(parts["docProps/core.xml"])),
		normalized(core.replace(" Alice ", "  Bob  ")),
	);
	for (const [name, bytes] of Object.entries(files)) {
		if (name !== "docProps/core.xml") assert.deepEqual(parts[name], bytes);
	}
});

test("new section properties keep schema order without moving content", async (t) => {
	const { input, output, files } = await fixture(t);
	const source = document
		.replace('<x:pgMar x:top="100"/>', "")
		.replace('<x:pgNumType x:start="1"/>', '<x:cols x:num="2"/>');
	files["word/document.xml"] = encode(source);
	await writeFile(input, zipSync(files));
	await patchDocx(input, output, { margins: { top: 720 }, pageNumbering: { start: 5 } });
	const expected = source
		.replace("<x:lnNumType", '<x:pgMar x:top="720"/><x:lnNumType')
		.replace("<x:cols", '<x:pgNumType x:start="5"/><x:cols');
	assert.deepEqual(
		normalized(decode(unzipSync(await readFile(output))["word/document.xml"])),
		normalized(expected),
	);
});

test("adding namespace declarations never rebinds an existing foreign prefix", async (t) => {
	const { input, output, files } = await fixture(t);
	const source = `<document xmlns="${word}" xmlns:w="urn:extension"><body><p><w:foreign w:value="001.230"/><r><t>unchanged</t></r></p></body></document>`;
	files["word/document.xml"] = encode(source);
	await writeFile(input, zipSync(files));
	await patchDocx(input, output, { page: { width: 1000 } });
	const xml = decode(unzipSync(await readFile(output))["word/document.xml"]);
	const expected = source
		.replace('xmlns:w="urn:extension"', `xmlns:w="urn:extension" xmlns:w1="${word}"`)
		.replace("</body>", '<w1:sectPr><w1:pgSz w1:w="1000"/></w1:sectPr></body>');
	assert.deepEqual(normalized(xml), normalized(expected));
	const metadata =
		'<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="urn:extension"><dc:foreign>untouched</dc:foreign></coreProperties>';
	files["docProps/core.xml"] = encode(metadata);
	await writeFile(input, zipSync(files));
	await patchDocx(input, output, { metadata: { creator: "Alice" } });
	const actual = decode(unzipSync(await readFile(output))["docProps/core.xml"]);
	assert.match(actual, /xmlns:dc="urn:extension"/);
	assert.match(actual, /xmlns:dc1="http:\/\/purl.org\/dc\/elements\/1.1\/"/);
	assert.match(actual, /<dc1:creator>Alice<\/dc1:creator>/);
	assert.match(actual, /<dc:foreign>untouched<\/dc:foreign>/);
});

test("formatting counts resolve custom comment relationships", async (t) => {
	const { input, files } = await fixture(t);
	files["word/review/notes.xml"] = encode(
		`<x:comments xmlns:x="${word}"><x:comment x:id="2"><x:p><x:ins/><x:del/></x:p></x:comment></x:comments>`,
	);
	delete files["word/comments.xml"];
	files["word/_rels/document.xml.rels"] = encode(
		'<Relationships><Relationship Id="notes" Type="urn:relationships/comments" Target="review/notes.xml"/></Relationships>',
	);
	await writeFile(input, zipSync(files));
	const info = await inspectDocx(input);
	assert.equal(info.comments, 1);
	assert.equal(info.trackedChanges, 4);
});

test("dry runs and errors protect source and existing output", async (t) => {
	const { input, output } = await fixture(t);
	const source = await readFile(input);
	await writeFile(output, "previous");
	const result = await patchDocx(input, output, { margins: { top: 500 } }, true);
	assert.equal(result.written, false);
	assert.equal(await readFile(output, "utf8"), "previous");
	await assert.rejects(() =>
		patchDocx(input, output, { lineNumbering: { mode: "off", start: 1 } }),
	);
	await assert.rejects(() => patchDocx(input, input, { margins: { top: 1 } }));
	assert.equal(await readFile(output, "utf8"), "previous");
	assert.deepEqual(await readFile(input), source);
});
