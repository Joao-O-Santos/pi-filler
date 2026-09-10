import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import extension, { executeFiller, normalizePath } from "../src/index.js";
import { readPdfText, renderPdfPages, searchPdf } from "../src/pdf.js";
import { MissingExecutableError, runCommand } from "../src/process.js";

async function fakeCommands(): Promise<{ directory: string; env: NodeJS.ProcessEnv }> {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-test-"));
	const script = async (name: string, body: string) => {
		const path = join(directory, name);
		await writeFile(path, `#!/bin/sh\n${body}\n`);
		await chmod(path, 0o755);
	};
	await script("pdftotext", 'printf "first page\\nsecond page\\n"');
	await script("pdfgrep", 'printf "fixture.pdf:2:Needle found\\n"');
	await script(
		"pdftocairo",
		'for arg do prefix="$arg"; done\ntouch "$prefix-1.png" "$prefix-2.png"',
	);
	return { directory, env: { ...process.env, PATH: `${directory}:${process.env.PATH ?? ""}` } };
}

test("extension registers one filler tool", () => {
	let registered: { name?: string } | undefined;
	extension({
		registerTool(definition: { name?: string }) {
			registered = definition;
		},
	} as never);
	assert.equal(registered?.name, "filler");
});

test("reports missing executables clearly", async () => {
	await assert.rejects(
		() => runCommand("pi-filler-command-that-does-not-exist", []),
		(error: unknown) =>
			error instanceof MissingExecutableError && /available on PATH/.test(error.message),
	);
});

test("normalizes relative and @ paths", () => {
	assert.equal(normalizePath("@docs/input.pdf", "/work/project"), "/work/project/docs/input.pdf");
	assert.equal(normalizePath("/tmp/input.pdf", "/work/project"), "/tmp/input.pdf");
});

test("runs deterministic PDF commands", async () => {
	const { directory, env } = await fakeCommands();
	const input = join(directory, "fixture.pdf");
	await writeFile(input, "not a real PDF");

	const text = await readPdfText(input, { firstPage: 2, lastPage: 2 }, { env });
	assert.equal(text.text, "first page\nsecond page\n");

	const search = await searchPdf(input, "Needle", { env });
	assert.deepEqual(search.matches, [{ page: 2, text: "Needle found" }]);
	assert.equal(search.text, "page 2: Needle found");

	const images = await renderPdfPages(input, join(directory, "render"), {}, { env });
	assert.deepEqual(images, [join(directory, "render-1.png"), join(directory, "render-2.png")]);
	assert.equal(await readFile(images[0], "utf8"), "");
});

test("dispatches PDF operations and rejects unsupported combinations", async () => {
	const { directory } = await fakeCommands();
	const input = join(directory, "fixture.pdf");
	await writeFile(input, "not a real PDF");
	await assert.rejects(
		() => executeFiller({ format: "pdf", action: "write", path: input }, directory),
		/PDF write and patch operations are not supported/,
	);
	await assert.rejects(
		() => executeFiller({ format: "pdf", action: "read", path: input }, directory),
		/PDF read requires view=text or view=image/,
	);
});
