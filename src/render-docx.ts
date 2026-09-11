import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, extname, join, resolve } from "node:path";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";
import { type PageRange, renderPdfPages } from "./pdf.js";

export async function renderDocxPages(
	path: string,
	output: string,
	range: PageRange = {},
	options?: RunCommandOptions,
): Promise<string[]> {
	const source = resolve(path);
	const temporaryDirectory = await mkdtemp(join(tmpdir(), "pi-filler-docx-"));
	try {
		const result = await runCommand(
			"libreoffice",
			["--headless", "--convert-to", "pdf", "--outdir", temporaryDirectory, source],
			{ ...options, cwd: options?.cwd ?? temporaryDirectory },
		);
		if (result.code !== 0) throw commandError(result, "libreoffice");
		const pdf = join(temporaryDirectory, `${basename(source, extname(source))}.pdf`);
		return await renderPdfPages(pdf, output, range, options);
	} finally {
		await rm(temporaryDirectory, { recursive: true, force: true });
	}
}
