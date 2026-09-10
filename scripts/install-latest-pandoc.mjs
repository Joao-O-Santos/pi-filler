#!/usr/bin/env node

import { execFileSync } from "node:child_process";
import { writeFile } from "node:fs/promises";

const releaseUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest";

const metadataResponse = await fetch(releaseUrl, {
	headers: { Accept: "application/vnd.github+json" },
});
if (!metadataResponse.ok) {
	throw new Error(`Pandoc metadata failed with HTTP ${metadataResponse.status}`);
}

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

const download = await fetch(asset.browser_download_url);
if (!download.ok) throw new Error(`Pandoc download failed with HTTP ${download.status}`);

const path = "/tmp/pandoc.deb";
await writeFile(path, Buffer.from(await download.arrayBuffer()));
execFileSync("dpkg", ["--install", path], { stdio: "inherit" });
