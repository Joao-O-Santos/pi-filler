import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { PDFDocument, StandardFonts } from "@cantoo/pdf-lib";
import { searchPdf } from "../src/pdf.js";
import { runCommand } from "../src/process.js";

// Exercises the real pdfgrep binary when it is installed; CI installs it.
async function pdfgrepAvailable(): Promise<boolean> {
	return runCommand("pdfgrep", ["--version"]).then(
		() => true,
		() => false,
	);
}

test("real pdfgrep reports page-numbered matches", async (t) => {
	if (!(await pdfgrepAvailable())) return t.skip("pdfgrep is not installed");
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-pdfgrep-"));
	try {
		const doc = await PDFDocument.create();
		const font = await doc.embedFont(StandardFonts.Helvetica);
		for (const line of ["Quarterly Revenue report", "Nothing relevant here"]) {
			doc.addPage([612, 792]).drawText(line, { x: 72, y: 700, size: 14, font });
		}
		const input = join(directory, "report.pdf");
		await writeFile(input, await doc.save());

		const found = await searchPdf(input, "revenue", { ignoreCase: true });
		assert.equal(found.matchCount, 1);
		assert.match(found.text, /^page 1: .*Revenue/m);

		const none = await searchPdf(input, "absent phrase");
		assert.equal(none.matchCount, 0);

		const literal = await searchPdf(input, "Quarterly.Revenue", { literal: true });
		assert.equal(literal.matchCount, 0, "fixed-strings treats the dot literally");
	} finally {
		await rm(directory, { recursive: true, force: true });
	}
});
