import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { executeFiller } from "../src/index.js";

async function script(directory: string, name: string, body: string): Promise<void> {
	const path = join(directory, name);
	await writeFile(path, `#!/bin/sh\n${body}\n`);
	await chmod(path, 0o755);
}

test("renders selected DOCX pages through LibreOffice and pdftocairo", async () => {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-docx-image-test-"));
	const input = join(directory, "input.docx");
	await writeFile(input, "fixture");
	await script(
		directory,
		"libreoffice",
		'printf "%s\\n" "$@" > "$LIBREOFFICE_LOG"\nwhile [ "$1" ]; do if [ "$1" = "--outdir" ]; then shift; out="$1"; fi; shift; done\ntouch "$out/input.pdf"',
	);
	await script(
		directory,
		"pdftocairo",
		'printf "%s\\n" "$@" > "$PDFTOCAIRO_LOG"\nfor arg do prefix="$arg"; done\ntouch "$prefix-2.png"',
	);
	const previous = process.env.PATH;
	process.env.PATH = `${directory}:${previous ?? ""}`;
	process.env.LIBREOFFICE_LOG = join(directory, "libreoffice-args");
	process.env.PDFTOCAIRO_LOG = join(directory, "pdftocairo-args");
	try {
		const result = await executeFiller(
			{
				format: "docx",
				action: "read",
				view: "image",
				path: input,
				output: join(directory, "render"),
				first_page: 2,
				last_page: 2,
			},
			directory,
		);
		assert.deepEqual(result.details.files, [join(directory, "render-2.png")]);
		assert.match(
			await readFile(process.env.LIBREOFFICE_LOG, "utf8"),
			/--headless\n--convert-to\npdf\n--outdir/,
		);
		assert.match(await readFile(process.env.PDFTOCAIRO_LOG, "utf8"), /-f\n2\n-l\n2/);
	} finally {
		process.env.PATH = previous;
		delete process.env.LIBREOFFICE_LOG;
		delete process.env.PDFTOCAIRO_LOG;
	}
});

test("DOCX image reads require an output path", async () => {
	await assert.rejects(
		() =>
			executeFiller({ format: "docx", action: "read", view: "image", path: "input.docx" }, "/tmp"),
		/DOCX image reads require output/,
	);
});
