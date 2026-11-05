#!/usr/bin/env node

import { execFileSync } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout } from "node:timers/promises";

async function fetchOk(url, label, options = {}) {
	for (let attempt = 0; ; attempt++) {
		try {
			const response = await fetch(url, { ...options, signal: AbortSignal.timeout(120_000) });
			if (response.ok) return response;
			await response.body?.cancel();
			const error = new Error(`Pandoc ${label} failed with HTTP ${response.status}`);
			if (![500, 502, 503, 504].includes(response.status)) throw error;
			if (attempt === 2) throw error;
			console.warn(`${error.message}; retrying`);
		} catch (error) {
			if (
				!(error instanceof TypeError) &&
				!(error instanceof Error && error.name === "TimeoutError")
			)
				throw error;
			if (attempt === 2) throw error;
			console.warn(`Pandoc ${label} request failed; retrying`);
		}
		await setTimeout(1000 * (attempt + 1));
	}
}

const releaseUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest";

const metadataResponse = await fetchOk(releaseUrl, "metadata", {
	headers: { Accept: "application/vnd.github+json" },
});

const metadata = await metadataResponse.json();
if (metadata.draft || metadata.prerelease || typeof metadata.tag_name !== "string") {
	throw new Error("Latest Pandoc metadata is not a stable release");
}

const assetName = `pandoc-${metadata.tag_name}-1-amd64.deb`;
const asset = metadata.assets?.find(
	(candidate) =>
		candidate?.name === assetName && typeof candidate.browser_download_url === "string",
);
if (!asset) throw new Error(`Could not find Pandoc asset ${assetName}`);

const download = await fetchOk(asset.browser_download_url, "download");
const directory = await mkdtemp(join(tmpdir(), "pi-filler-pandoc-"));
try {
	const path = join(directory, "pandoc.deb");
	await writeFile(path, Buffer.from(await download.arrayBuffer()));
	execFileSync("dpkg", ["--install", path], { stdio: "inherit" });
} finally {
	await rm(directory, { recursive: true, force: true });
}
