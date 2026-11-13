import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import extension, { executeFiller, type FillerInput, fillerSchema } from "../src/index.js";

const old = "PRIVATE old sentence 7c9";
const replacement = "PRIVATE new sentence 8d0";
const author = "PRIVATE author 9e1";
const digest = createHash("sha256").update("source bytes").digest("hex");

async function fixture(t: TestContext, mode = "success") {
	const dir = await mkdtemp(join(tmpdir(), "filler-toolkit-"));
	const previous = { ...process.env };
	t.after(async () => {
		process.env.PATH = previous.PATH;
		delete process.env.TOOLKIT_TEST_LOG;
		delete process.env.TOOLKIT_TEST_MODE;
		await rm(dir, { recursive: true, force: true });
	});
	process.env.PATH = dir;
	process.env.TOOLKIT_TEST_LOG = join(dir, "log");
	process.env.TOOLKIT_TEST_MODE = mode;
	const executable = join(dir, "docxquisite");
	await writeFile(
		executable,
		`#!${process.execPath}
const fs = require('node:fs');
let raw = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => raw += chunk);
process.stdin.on('end', () => {
  const request = JSON.parse(raw);
  fs.appendFileSync(process.env.TOOLKIT_TEST_LOG, JSON.stringify({argv: process.argv.slice(2), request}) + '\\n');
  const mode = process.env.TOOLKIT_TEST_MODE;
  if (mode === 'malformed') { console.log('{broken'); return; }
  if (mode === 'exit') { console.error('PRIVATE old sentence 7c9 PRIVATE new sentence 8d0 PRIVATE author 9e1'); process.exitCode = 1; return; }
  if (mode === 'error') { console.log(JSON.stringify({version:1,ok:false,error:{code:'document_error',message:'PRIVATE old sentence 7c9 PRIVATE new sentence 8d0 PRIVATE author 9e1'}})); process.exitCode = 1; return; }
  if (request.operation === 'capabilities') {
    console.log(JSON.stringify({version:1,ok:true,result:{protocol_version:mode === 'version' ? 2 : 1,operations:mode === 'unsupported' ? ['capabilities'] : ['capabilities','inspect','audit','replace_tracked'],replace_limitations:[]}}));
    return;
  }
  if (mode === 'replace-malformed') { console.log('not json'); return; }
  if (mode === 'replace-error') { console.error('PRIVATE old sentence 7c9 PRIVATE new sentence 8d0 PRIVATE author 9e1'); console.log(JSON.stringify({version:1,ok:false,error:{code:'document_error',message:'PRIVATE old sentence 7c9 PRIVATE new sentence 8d0 PRIVATE author 9e1'}})); process.exitCode = 1; return; }
  // docxquisite replaces an existing distinct output and, with overwrite_original, the source.
  // It hashes the source it read before publishing, so the hash is taken first.
  const sourceSha = require('node:crypto').createHash('sha256').update(fs.readFileSync(request.input)).digest('hex');
  fs.writeFileSync(request.output, 'tracked output');
  console.log(JSON.stringify({version:1,ok:true,result:{output:request.output,source_sha256:sourceSha,output_sha256:require('node:crypto').createHash('sha256').update('tracked output').digest('hex'),checks:[],limitations:[]}}));
});
`,
	);
	await chmod(executable, 0o755);
	await writeFile(join(dir, "source.docx"), "source bytes");
	const input = {
		format: "docx",
		action: "patch",
		path: "@source.docx",
		output: "@output.docx",
		tracked_replace: { old, new: replacement, author, expected_sha256: digest },
	};
	return {
		dir,
		executable,
		input,
		source: join(dir, "source.docx"),
		output: join(dir, "output.docx"),
	};
}

async function call(input: object, dir: string) {
	return executeFiller(input as FillerInput, dir);
}

async function log(dir: string) {
	try {
		return (await readFile(join(dir, "log"), "utf8"))
			.trim()
			.split("\n")
			.map((line) => JSON.parse(line));
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
		throw error;
	}
}

async function absent(path: string) {
	await assert.rejects(readFile(path), { code: "ENOENT" });
}

test("tracked replacement sends protocol requests over stdin, resolves paths and preserves source", async (t) => {
	const f = await fixture(t);
	const result = await call(f.input, f.dir);
	assert.ok(result);
	const calls = await log(f.dir);
	assert.equal(calls.length, 2);
	assert.deepEqual(
		calls.map((entry) => entry.argv),
		[["protocol"], ["protocol"]],
	);
	assert.deepEqual(calls[0].request, { version: 1, operation: "capabilities" });
	assert.deepEqual(calls[1].request, {
		version: 1,
		operation: "replace_tracked",
		input: f.source,
		output: f.output,
		old,
		new: replacement,
		author,
		expected_sha256: digest,
	});
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
	assert.equal(await readFile(f.output, "utf8"), "tracked output");
});

test("tracked deletion allows empty new and binds the source hash", async (t) => {
	const f = await fixture(t);
	await call({ ...f.input, tracked_replace: { old, new: "", author } }, f.dir);
	const calls = await log(f.dir);
	assert.equal(calls[1].request.new, "");
	assert.equal(calls[1].request.expected_sha256, digest);
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
	assert.equal(await readFile(f.output, "utf8"), "tracked output");
});

test("invalid tracked requests fail before invoking the executable", async (t) => {
	const f = await fixture(t);
	const invalid = [
		{ tracked_replace: { old: "", new: replacement, author } },
		{ tracked_replace: { old, new: replacement, author: "" } },
		{ tracked_replace: { old, author } },
		{ tracked_replace: { old, new: replacement, author, expected_sha256: "A".repeat(64) } },
		{ tracked_replace: { old, new: replacement, author, expected_sha256: "a".repeat(63) } },
		{ tracked_replace: { old, new: replacement, author, expected_sha256: "g".repeat(64) } },
		{ patches: { clearCoreMetadata: true } },
		{ dry_run: true },
		{ dry_run: false },
		{ view: "text" },
		{ query: "irrelevant" },
		{ output: "@source.docx" },
	];
	for (const change of invalid) {
		await assert.rejects(call({ ...f.input, ...change }, f.dir), JSON.stringify(change));
	}
	assert.deepEqual(await log(f.dir), []);
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
	await absent(f.output);
});

test("missing executable fails safely", async (t) => {
	const f = await fixture(t);
	await rm(f.executable);
	await assert.rejects(call(f.input, f.dir));
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
	await absent(f.output);
});

for (const mode of [
	"version",
	"unsupported",
	"malformed",
	"exit",
	"error",
	"replace-malformed",
	"replace-error",
]) {
	test(`protocol ${mode} fails without exposing request secrets`, async (t) => {
		const f = await fixture(t, mode);
		let failure: unknown;
		try {
			await call(f.input, f.dir);
			assert.fail("protocol failure must reject");
		} catch (error) {
			failure = error;
		}
		assert.ok(failure instanceof Error);
		assert.notEqual(failure.message, "protocol failure must reject");
		const exposed = `${failure.stack} ${JSON.stringify(failure, Object.getOwnPropertyNames(failure))}`;
		for (const secret of [old, replacement, author]) assert.ok(!exposed.includes(secret));
		const calls = await log(f.dir);
		assert.equal(calls.length, mode.startsWith("replace-") ? 2 : 1);
		assert.equal(await readFile(f.source, "utf8"), "source bytes");
		await absent(f.output);
	});
}

test("existing distinct output is replaced and the source stays unchanged", async (t) => {
	const f = await fixture(t);
	await writeFile(f.output, "existing output");
	await call(f.input, f.dir);
	assert.equal(await readFile(f.output, "utf8"), "tracked output");
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
});

test("schema and prompt expose tracked replacement", () => {
	const schema = JSON.parse(JSON.stringify(fillerSchema));
	assert.ok(schema.properties.tracked_replace);
	assert.match(schema.properties.tracked_replace.description, /docxquisite on PATH/);
	let registered: { promptSnippet?: string; promptGuidelines?: string[] } | undefined;
	extension({
		registerTool(tool: typeof registered) {
			registered = tool;
		},
	} as never);
	assert.match(
		`${registered?.promptSnippet} ${registered?.promptGuidelines?.join(" ")}`,
		/tracked_replace/,
	);
	assert.match(registered?.promptGuidelines?.join(" ") ?? "", /docxquisite on PATH/);
});

test("tracked replacement replaces an existing output and forwards overwrite_original only when requested", async (t) => {
	const f = await fixture(t);
	await writeFile(f.output, "previous output");
	await call(f.input, f.dir);
	assert.equal(await readFile(f.output, "utf8"), "tracked output");
	assert.deepEqual((await log(f.dir))[1].request.overwrite_original, undefined);
});

test("tracked replacement can replace the source in place only with overwrite_original", async (t) => {
	const f = await fixture(t);
	const inPlace = { ...f.input, output: "@source.docx" };
	await assert.rejects(call(inPlace, f.dir), /overwrite_original/);
	assert.equal(await readFile(f.source, "utf8"), "source bytes");
	assert.equal((await log(f.dir)).length, 0);
	await call({ ...inPlace, overwrite_original: true }, f.dir);
	assert.equal(await readFile(f.source, "utf8"), "tracked output");
	const calls = await log(f.dir);
	assert.equal(calls[calls.length - 1].request.overwrite_original, true);
});
