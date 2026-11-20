import { randomUUID } from "node:crypto";
import { copyFile, mkdir, mkdtemp, readFile, rename, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { PDFDocument } from "@cantoo/pdf-lib";
import { distinct } from "./docx-package.js";
import { commandError, type RunCommandOptions, runCommand } from "./process.js";

/** LibreOffice is a renderer, not proof of Microsoft Office layout equivalence. */
export async function exportOfficePdf(
	inputPath: string,
	outputPath: string,
	options: RunCommandOptions = {},
) {
	const input = resolve(inputPath),
		output = resolve(outputPath);
	const extension = extname(input).toLowerCase();
	if (![".docx", ".pptx", ".xlsx"].includes(extension) || extname(output).toLowerCase() !== ".pdf")
		throw new Error("PDF export requires DOCX/PPTX/XLSX input and a .pdf output");
	await distinct(output, [input]);
	const temporary = await mkdtemp(join(tmpdir(), "pi-filler-export-"));
	let sibling: string | undefined;
	try {
		const version = await runCommand("libreoffice", ["--version"], options);
		if (version.code !== 0) throw commandError(version, "libreoffice");
		const backendVersion = version.stdout.toString("utf8").trim().slice(0, 200);
		if (!backendVersion) throw new Error("LibreOffice returned no backend version");
		const staged = join(temporary, `source${extension}`);
		await copyFile(input, staged);
		const result = await runCommand(
			"libreoffice",
			[
				`-env:UserInstallation=${pathToFileURL(join(temporary, "profile")).href}`,
				"--headless",
				"--convert-to",
				"pdf",
				"--outdir",
				temporary,
				staged,
			],
			{ ...options, cwd: temporary },
		);
		if (result.code !== 0) throw commandError(result, "libreoffice");
		// A zero exit status alone does not establish that a PDF was produced.
		const bytes = await readFile(join(temporary, "source.pdf"));
		const pdf = await PDFDocument.load(bytes, { updateMetadata: false });
		const pageCount = pdf.getPageCount();
		if (pageCount < 1) throw new Error("Export produced an empty PDF");
		options.signal?.throwIfAborted();
		await mkdir(dirname(output), { recursive: true });
		sibling = join(dirname(output), `.pi-filler-${randomUUID()}.pdf`);
		await writeFile(sibling, bytes);
		options.signal?.throwIfAborted();
		await rename(sibling, output);
		return {
			output,
			pageCount,
			backend: { name: "LibreOffice", version: backendVersion },
			warnings: [
				"LibreOffice rendering is not Microsoft Office layout equivalence. Page count does not detect orphaned blocks or overflow; inspect page images.",
			],
		};
	} finally {
		if (sibling) await rm(sibling, { force: true });
		await rm(temporary, { recursive: true, force: true });
	}
}
