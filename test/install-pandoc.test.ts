import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { access, chmod, mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const installer = fileURLToPath(new URL("../scripts/install-latest-pandoc.mjs", import.meta.url));
const preload = `
const steps = JSON.parse(process.env.PANDOC_TEST_STEPS);
globalThis.fetch = async (url, options) => {
 const metadata = url === 'https://api.github.com/repos/jgm/pandoc/releases/latest';
 if (metadata && options.headers.Accept !== 'application/vnd.github+json') throw Error('missing metadata header');
 const step = steps.shift();
 console.log('REQUEST',metadata ? 'metadata' : 'asset',step);
 if (step === undefined) throw Error('unexpected extra request');
 if (step === 'network') throw new TypeError('fixture transport failure');
 if (step === 'body-fail') return new Response(new ReadableStream({start(c){c.error(Error('fixture truncated body'));}}));
 if (step !== 200) return new Response('fixture HTTP failure',{status:step});
 return new Response(metadata ? JSON.stringify({tag_name:'9.8.7',draft:false,prerelease:false,assets:[{name:'pandoc-9.8.7-1-amd64.deb',browser_download_url:'https://example.invalid/pandoc.deb'}]}) : 'fixture deb bytes');
};
`;

const cases = [
	{
		name: "retries a transient asset failure",
		steps: [200, 504, 200],
		success: true,
		install: true,
	},
	{
		name: "retries transient metadata failure",
		steps: [503, 200, 200],
		success: true,
		install: true,
	},
	{ name: "retries transport failure", steps: ["network", 200, 200], success: true, install: true },
	{
		name: "stops after three failed asset attempts",
		steps: [200, 504, 504, 504],
		error: /download failed with HTTP 504/,
		success: false,
		install: false,
	},
	{
		name: "does not retry a missing release",
		steps: [404],
		error: /metadata failed with HTTP 404/,
		success: false,
		install: false,
	},
	{
		name: "does not retry rate limits",
		steps: [200, 429],
		error: /download failed with HTTP 429/,
		success: false,
		install: false,
	},
	{
		name: "cleans downloaded files when installation fails",
		steps: [200, 200],
		error: /fixture dpkg failure/,
		success: false,
		install: true,
		dpkgFail: true,
	},
	{
		name: "cleans temporary files on truncated response bodies",
		steps: [200, "body-fail"],
		error: /fixture truncated body/,
		success: false,
		install: false,
	},
];

for (const fixture of cases) {
	test(`Pandoc installer ${fixture.name}`, async (t) => {
		const directory = await mkdtemp(join(tmpdir(), "pi-filler-installer-test-"));
		t.after(() => rm(directory, { recursive: true, force: true }));
		const scratch = join(directory, "scratch");
		await mkdir(scratch);
		const log = join(directory, "dpkg-path");
		const archive = join(directory, "installed-bytes");
		const dpkg = join(directory, "dpkg");
		await writeFile(
			dpkg,
			`#!/bin/sh\n[ "$1" = "--install" ] || exit 21\n[ -f "$2" ] || exit 22\nprintf '%s' "$2" > "$PANDOC_TEST_LOG"\ncat "$2" > "$PANDOC_TEST_ARCHIVE"\nif [ "$PANDOC_TEST_DPKG_FAIL" = "1" ]; then echo 'fixture dpkg failure' >&2; exit 13; fi\n`,
		);
		await chmod(dpkg, 0o755);
		let stdout = "";
		let stderr = "";
		let success = true;
		try {
			const result = await promisify(execFile)(
				process.execPath,
				["--import", `data:text/javascript,${encodeURIComponent(preload)}`, installer],
				{
					env: {
						...process.env,
						PATH: `${directory}:${process.env.PATH ?? ""}`,
						TMPDIR: scratch,
						PANDOC_TEST_STEPS: JSON.stringify(fixture.steps),
						PANDOC_TEST_LOG: log,
						PANDOC_TEST_ARCHIVE: archive,
						PANDOC_TEST_DPKG_FAIL: fixture.dpkgFail ? "1" : "0",
					},
					timeout: 15000,
				},
			);
			stdout = result.stdout;
			stderr = result.stderr;
		} catch (error) {
			const result = error as Error & { stdout: string; stderr: string };
			success = false;
			stdout = result.stdout;
			stderr = result.stderr;
		}
		assert.equal(success, fixture.success, stderr);
		assert.equal(
			stdout.split("\n").filter((line) => line.startsWith("REQUEST ")).length,
			fixture.steps.length,
		);
		if (fixture.error) assert.match(stderr, fixture.error);
		if (fixture.install) {
			assert.equal(await readFile(archive, "utf8"), "fixture deb bytes");
			const packagePath = await readFile(log, "utf8");
			assert.ok(packagePath.startsWith(join(scratch, "pi-filler-pandoc-")));
			await assert.rejects(access(packagePath), { code: "ENOENT" });
		} else await assert.rejects(access(log), { code: "ENOENT" });
		assert.deepEqual(await readdir(scratch), []);
	});
}
