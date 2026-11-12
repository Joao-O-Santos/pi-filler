import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { searchPdf } from "../src/pdf.js";

test("PDF search passes open and single-page ranges and handles no matches", async (t) => {
	const dir = await mkdtemp(join(tmpdir(), "pdf-search-extra-"));
	const previous = process.env.PATH;
	process.env.PATH = dir;
	t.after(async () => {
		if (previous === undefined) delete process.env.PATH;
		else process.env.PATH = previous;
		await rm(dir, { recursive: true, force: true });
	});
	await writeFile(
		join(dir, "pdfgrep"),
		`#!${process.execPath}
const fs = require('node:fs');
fs.writeFileSync('args.json', JSON.stringify(process.argv.slice(2)));
process.exitCode = 1;
`,
	);
	await chmod(join(dir, "pdfgrep"), 0o755);
	for (const [range, expected] of [
		[{ firstPage: 3 }, "3-2147483647"],
		[{ lastPage: 4 }, "1-4"],
		[{ firstPage: 2, lastPage: 2 }, "2"],
	] as const) {
		const result = await searchPdf(join(dir, "source.pdf"), "needle", { ...range, cwd: dir });
		assert.deepEqual(result, { matchCount: 0, text: "" });
		const args: string[] = JSON.parse(await readFile(join(dir, "args.json"), "utf8"));
		assert.equal(args[args.indexOf("--page-range") + 1], expected);
		assert.deepEqual(args.slice(-3), ["--", "needle", join(dir, "source.pdf")]);
	}
});
