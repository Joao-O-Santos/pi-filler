import assert from "node:assert/strict";
import { access, chmod, mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import test, { type TestContext } from "node:test";
import { executeFiller } from "../src/index.js";

async function fixture(t: TestContext, mode = "success") {
	const directory = await mkdtemp(join(tmpdir(), "pi-filler-profile-test-"));
	const scratch = join(directory, "scratch with spaces");
	await mkdir(scratch);
	const log = join(directory, "profiles.jsonl");
	const changes = {
		PATH: `${directory}:${process.env.PATH ?? ""}`,
		TMPDIR: scratch,
		OFFICE_PROFILE_LOG: log,
		OFFICE_TEST_MODE: mode,
	};
	const previous = Object.fromEntries(Object.keys(changes).map((key) => [key, process.env[key]]));
	t.after(async () => {
		for (const [key, value] of Object.entries(previous)) {
			if (value === undefined) delete process.env[key];
			else process.env[key] = value;
		}
		await rm(directory, { recursive: true, force: true });
	});
	for (const [name, code] of Object.entries({
		libreoffice: `
const fs = require('node:fs');
const path = require('node:path');
const {fileURLToPath} = require('node:url');
const args = process.argv.slice(2);
const flag = args.find(a => a.startsWith('-env:UserInstallation='));
if (!flag) throw Error('missing isolated profile');
const uri = flag.slice('-env:UserInstallation='.length);
const profile = fileURLToPath(uri);
const out = args[args.indexOf('--outdir') + 1];
if (path.dirname(profile) !== out || path.basename(profile) !== 'profile') throw Error('profile not private');
fs.mkdirSync(profile);
fs.appendFileSync(process.env.OFFICE_PROFILE_LOG, JSON.stringify({uri,profile,out})+'\\n');
if (process.env.OFFICE_TEST_MODE === 'converter-fail') {console.error('converter fixture failure'); process.exit(44);}
const input = args.at(-1);
if (!fs.existsSync(input)) throw Error('missing source');
fs.writeFileSync(path.join(out, path.basename(input, path.extname(input))+'.pdf'), 'temporary PDF');
`,
		pdftocairo: `
const fs = require('node:fs');
const args = process.argv.slice(2);
const pdf = args.find(a => a.endsWith('.pdf'));
if (!pdf || fs.readFileSync(pdf,'utf8') !== 'temporary PDF') throw Error('missing converted PDF');
if (process.env.OFFICE_TEST_MODE === 'render-fail') {console.error('renderer fixture failure'); process.exit(45);}
fs.writeFileSync(args.at(-1)+'-1.png', 'page image');
`,
	})) {
		const executable = join(directory, name);
		await writeFile(executable, `#!/usr/bin/env node\n${code}`);
		await chmod(executable, 0o755);
	}
	Object.assign(process.env, changes);
	return { directory, scratch, log };
}

function input(format: "docx" | "pptx" | "xlsx", directory: string) {
	return {
		format,
		action: "read" as const,
		view: "image" as const,
		path: join(directory, `input.${format}`),
		output: join(directory, `${format}-render`),
	};
}

test("concurrent office renders have unique encoded profiles and clean all temporary data", async (t) => {
	const { directory, scratch, log } = await fixture(t);
	const formats = ["docx", "pptx", "xlsx"] as const;
	for (const format of formats) await writeFile(input(format, directory).path, "source");
	await Promise.all(
		formats.map(async (format) => {
			const result = await executeFiller(input(format, directory), directory);
			assert.deepEqual(result.details.files, [join(directory, `${format}-render-1.png`)]);
			assert.equal(await readFile(input(format, directory).path, "utf8"), "source");
		}),
	);
	const records = (await readFile(log, "utf8"))
		.trim()
		.split("\n")
		.map((line) => JSON.parse(line));
	assert.equal(records.length, 3);
	assert.equal(new Set(records.map((record) => record.profile)).size, 3);
	for (const record of records) {
		assert.match(record.uri, /^file:\/\//);
		assert.match(record.uri, /%20/);
		await assert.rejects(access(record.profile), { code: "ENOENT" });
		await assert.rejects(access(dirname(record.profile)), { code: "ENOENT" });
	}
	assert.deepEqual(await readdir(scratch), []);
});

for (const mode of ["converter-fail", "render-fail"] as const) {
	test(`${mode} cleans profiles and preserves the actual target's prior render`, async (t) => {
		const { directory, scratch, log } = await fixture(t, mode);
		const request = input("docx", directory);
		await writeFile(request.path, "unchanged source");
		const image = `${request.output}-1.png`;
		await writeFile(image, "previous image");
		await assert.rejects(executeFiller(request, directory), /fixture failure/);
		assert.equal(await readFile(image, "utf8"), "previous image");
		assert.equal(await readFile(request.path, "utf8"), "unchanged source");
		const record = JSON.parse((await readFile(log, "utf8")).trim());
		await assert.rejects(access(record.profile), { code: "ENOENT" });
		await assert.rejects(access(record.out), { code: "ENOENT" });
		assert.deepEqual(await readdir(scratch), []);
	});
}

test("pre-aborted rendering preserves input/output and removes the private directory", async (t) => {
	const { directory, scratch } = await fixture(t);
	const request = input("xlsx", directory);
	await writeFile(request.path, "source");
	await writeFile(`${request.output}-1.png`, "old image");
	const controller = new AbortController();
	controller.abort();
	await assert.rejects(executeFiller(request, directory, controller.signal), /aborted/);
	assert.equal(await readFile(`${request.output}-1.png`, "utf8"), "old image");
	assert.equal(await readFile(request.path, "utf8"), "source");
	assert.deepEqual(await readdir(scratch), []);
});
