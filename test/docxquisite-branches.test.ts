import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { chmod, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { commandError, runCommand } from "../src/process.js";
import { replaceTracked, validateTrackedReplacement } from "../src/tracked-replace.js";

const hash = (value: string) => createHash("sha256").update(value).digest("hex");
const replacement = { old: "secret old", new: "secret new", author: "secret author" };

async function fixture(t: TestContext) {
	const dir = await mkdtemp(join(tmpdir(), "tracked-branches-"));
	const previous = process.env.PATH;
	process.env.PATH = dir;
	t.after(async () => {
		if (previous === undefined) delete process.env.PATH;
		else process.env.PATH = previous;
		await rm(dir, { recursive: true, force: true });
	});
	await writeFile(join(dir, "source"), "source");
	await writeFile(
		join(dir, "docxquisite"),
		`#!${process.execPath}
const fs = require('node:fs');
const path = require('node:path');
let input = '';
process.stdin.on('data', chunk => input += chunk);
process.stdin.on('end', () => {
  const request = JSON.parse(input);
  const config = JSON.parse(fs.readFileSync(path.join(process.cwd(), 'config')));
  fs.appendFileSync(path.join(process.cwd(), 'calls'), request.operation + '\\n');
  const caps = {version:1,ok:true,result:{protocol_version:1,operations:['replace_tracked'],replace_limitations:[]}};
  if (request.operation === 'capabilities') {
    process.stdout.write(config.capRaw ?? JSON.stringify('cap' in config ? config.cap : caps));
    process.exitCode = config.capCode ?? 0;
  } else {
    if (config.output === 'file') fs.writeFileSync(request.output, 'output');
    if (config.output === 'directory') fs.mkdirSync(request.output);
    const result = {output:request.output,source_sha256:${JSON.stringify(hash("source"))},output_sha256:${JSON.stringify(hash("output"))},checks:[],limitations:[]};
    process.stdout.write(config.raw ?? JSON.stringify(config.response ?? {version:1,ok:true,result:{...result,...config.result}}));
    process.exitCode = config.code ?? 0;
  }
});
`,
	);
	await chmod(join(dir, "docxquisite"), 0o755);
	return {
		dir,
		input: join(dir, "source"),
		output: join(dir, "output"),
		config: (value: object) => writeFile(join(dir, "config"), JSON.stringify(value)),
		calls: async () => {
			try {
				return (await readFile(join(dir, "calls"), "utf8")).trim().split("\n");
			} catch (error) {
				if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
				throw error;
			}
		},
	};
}

test("tracked replacement validates object shape and value types", () => {
	for (const value of [
		null,
		[],
		"text",
		{ ...replacement, extra: true },
		{ ...replacement, old: 1 },
		{ ...replacement, new: null },
		{ ...replacement, author: false },
		{ ...replacement, expected_sha256: 1 },
	]) {
		assert.throws(() => validateTrackedReplacement(value), /tracked_replace requires/);
	}
	validateTrackedReplacement({ ...replacement, new: "", expected_sha256: hash("source") });
});

test("capability envelopes and error statuses are strictly validated", async (t) => {
	const f = await fixture(t);
	const caps = { protocol_version: 1, operations: ["replace_tracked"], replace_limitations: [] };
	for (const config of [
		{ cap: null },
		{ cap: [] },
		{ cap: { version: 2, ok: true, result: caps } },
		{ cap: { version: 1, ok: "true", result: caps } },
		{ cap: { version: 1, ok: true, result: caps }, capCode: 1 },
		{ cap: { version: 1, ok: true, result: [], extra: 1 } },
		{ cap: { version: 1, ok: true, result: [] } },
		{ cap: { version: 1, ok: false, error: { code: "document_error", message: "secret" } } },
		{ cap: { version: 1, ok: false, error: null }, capCode: 1 },
		{ cap: { version: 1, ok: false, error: { code: 4, message: "x" } }, capCode: 1 },
		{ cap: { version: 1, ok: false, error: { code: "unknown", message: "x" } }, capCode: 1 },
		{ cap: { version: 1, ok: false, error: { code: "io_error", message: 4 } }, capCode: 1 },
		{
			cap: { version: 1, ok: false, error: { code: "io_error", message: "x", extra: 1 } },
			capCode: 1,
		},
		{ cap: { version: 1, ok: true, result: { ...caps, operations: [1] } } },
		{ cap: { version: 1, ok: true, result: { ...caps, replace_limitations: [1] } } },
		{ cap: { version: 1, ok: true, result: { ...caps, extra: 1 } } },
	]) {
		await f.config(config);
		await assert.rejects(replaceTracked(f.input, f.output, replacement, f.dir), /capability|probe/);
	}
	await f.config({
		cap: { version: 1, ok: false, error: { code: "io_error", message: "secret" } },
		capCode: 1,
	});
	await assert.rejects(
		replaceTracked(f.input, f.output, replacement, f.dir),
		/rejected.*io_error.*No replacement/,
	);
	assert.ok((await f.calls()).every((call) => call === "capabilities"));
});

test("replacement rejects untrusted result fields and preserves any output", async (t) => {
	const f = await fixture(t);
	for (const result of [
		{ output: "other" },
		{ source_sha256: "0".repeat(64) },
		{ output_sha256: 3 },
		{ output_sha256: "A".repeat(64) },
		{ checks: [1] },
		{ limitations: null },
		{ extra: true },
	]) {
		await f.config({ result, output: "file" });
		await assert.rejects(
			replaceTracked(f.input, f.output, replacement, f.dir),
			/outcome may be unknown/,
		);
		assert.equal(await readFile(f.output, "utf8"), "output");
		await rm(f.output);
	}
	await f.config({ output: "directory" });
	await assert.rejects(
		replaceTracked(f.input, f.output, replacement, f.dir),
		/outcome may be unknown/,
	);
	await rm(f.output, { recursive: true });
	await f.config({
		response: { version: 1, ok: false, error: { code: "output_exists", message: "secret" } },
		code: 1,
	});
	await assert.rejects(
		replaceTracked(f.input, f.output, replacement, f.dir),
		/rejected.*output_exists.*inspect/,
	);
});

test("fingerprint, size and transport failures stop safely", async (t) => {
	const f = await fixture(t);
	await f.config({});
	await assert.rejects(replaceTracked(f.input, f.input, replacement, f.dir), /must differ/);
	await assert.rejects(
		replaceTracked(join(f.dir, "missing"), f.output, replacement, f.dir),
		/fingerprint/,
	);
	await assert.rejects(
		replaceTracked(f.input, f.output, { ...replacement, expected_sha256: "0".repeat(64) }, f.dir),
		/SHA-256 mismatch/,
	);
	await assert.rejects(
		replaceTracked(f.input, f.output, { ...replacement, old: "x".repeat(1024 * 1024) }, f.dir),
		/outcome may be unknown/,
	);
	assert.deepEqual(await f.calls(), ["capabilities", "capabilities", "capabilities"]);
	await f.config({ raw: "x".repeat(70_000) });
	await assert.rejects(
		replaceTracked(f.input, f.output, replacement, f.dir),
		/outcome may be unknown/,
	);
	await f.config({ raw: "\\ud800" });
	await assert.rejects(
		replaceTracked(f.input, f.output, replacement, f.dir),
		/outcome may be unknown/,
	);
});

test("process adapter reports bounded output, timeout, abort, and exit diagnostics", async (t) => {
	const dir = await mkdtemp(join(tmpdir(), "process-branches-"));
	t.after(() => rm(dir, { recursive: true, force: true }));
	await mkdir(join(dir, "cwd"));
	await assert.rejects(
		runCommand(process.execPath, ["-e", "process.stdout.write('x'.repeat(10000))"], {
			maxBufferBytes: 100,
		}),
		/output exceeded 100 bytes/,
	);
	await assert.rejects(
		runCommand(process.execPath, ["-e", "setInterval(() => {}, 1000)"], { timeoutMs: 50 }),
		/timed out after 50ms/,
	);
	const controller = new AbortController();
	controller.abort();
	await assert.rejects(runCommand(process.execPath, [], { signal: controller.signal }), /aborted/);
	const result = await runCommand(process.execPath, [
		"-e",
		"console.error('diagnostic'); process.exitCode=7",
	]);
	assert.equal(result.code, 7);
	assert.equal(commandError(result, "tool").message, "diagnostic");
	assert.equal(
		commandError({ ...result, stderr: Buffer.alloc(0) }, "tool").message,
		"tool exited with code 7",
	);
	await assert.rejects(
		runCommand(process.execPath, [], { cwd: join(dir, "missing") }),
		/not found/,
	);
});
