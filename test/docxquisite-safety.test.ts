import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { executeFiller, type FillerInput } from "../src/index.js";

const hash = (value: string) => createHash("sha256").update(value).digest("hex");
const secret = "PRIVATE-stdin-secret-6b5e";

async function fixture(t: TestContext, mode: string) {
	const dir = await mkdtemp(join(tmpdir(), "filler-safety-"));
	const previous = { ...process.env };
	t.after(async () => {
		process.env.PATH = previous.PATH;
		for (const key of ["SAFETY_MODE", "SAFETY_DIR"]) {
			if (previous[key] === undefined) delete process.env[key];
			else process.env[key] = previous[key];
		}
		await rm(dir, { recursive: true, force: true });
	});
	process.env.PATH = dir;
	process.env.SAFETY_MODE = mode;
	process.env.SAFETY_DIR = dir;
	const executable = join(dir, "docxquisite");
	await writeFile(
		executable,
		`#!${process.execPath}
const fs = require('node:fs');
const crypto = require('node:crypto');
const path = require('node:path');
const dir = process.env.SAFETY_DIR;
const mode = process.env.SAFETY_MODE;
const marker = path.join(dir, 'capabilities-done');
const digest = s => crypto.createHash('sha256').update(s).digest('hex');
if (mode === 'closed-stdin' && fs.existsSync(marker)) {
  fs.closeSync(0);
  fs.writeFileSync(path.join(dir, 'closed'), 'yes');
  const output = path.join(dir, 'output.docx');
  fs.writeFileSync(output, 'tracked output');
  console.log(JSON.stringify({version:1,ok:true,result:{output,source_sha256:digest('source bytes'),output_sha256:digest('tracked output'),checks:[],limitations:[]}}));
} else {
  let raw = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', chunk => raw += chunk);
  process.stdin.on('end', () => {
    const request = JSON.parse(raw);
    if (request.operation === 'capabilities') {
      fs.writeFileSync(marker, 'yes');
      console.log(JSON.stringify({version:1,ok:true,result:{protocol_version:1,operations:['capabilities','replace_tracked'],replace_limitations:[]}}));
      return;
    }
    fs.writeFileSync(path.join(dir, 'request'), raw);
    if (mode !== 'missing-output') fs.writeFileSync(request.output, 'tracked output');
    console.log(JSON.stringify({version:1,ok:true,result:{output:request.output,source_sha256:digest('source bytes'),output_sha256:mode === 'wrong-hash' ? '0'.repeat(64) : digest('tracked output'),checks:[],limitations:[]}}));
  });
}
`,
	);
	await chmod(executable, 0o755);
	await writeFile(join(dir, "source.docx"), "source bytes");
	const input = {
		format: "docx",
		action: "patch",
		path: "@source.docx",
		output: "@output.docx",
		tracked_replace: { old: secret, new: "PRIVATE replacement", author: "PRIVATE author" },
	};
	return { dir, executable, input, output: join(dir, "output.docx") };
}

function call(input: object, dir: string) {
	return executeFiller(input as FillerInput, dir);
}

async function rejected(promise: Promise<unknown>) {
	let error: unknown;
	try {
		await promise;
	} catch (caught) {
		error = caught;
	}
	assert.ok(error instanceof Error, "operation must reject");
	return error;
}

test("omitted expected hash is computed from the source and sent upstream", async (t) => {
	const f = await fixture(t, "success");
	await call(f.input, f.dir);
	const request = JSON.parse(await readFile(join(f.dir, "request"), "utf8"));
	assert.equal(request.expected_sha256, hash("source bytes"));
});

for (const mode of ["missing-output", "wrong-hash"]) {
	test(`success response with ${mode} is rejected`, async (t) => {
		const f = await fixture(t, mode);
		await rejected(call(f.input, f.dir));
		assert.equal(await readFile(join(f.dir, "source.docx"), "utf8"), "source bytes");
		if (mode === "missing-output") await assert.rejects(readFile(f.output), { code: "ENOENT" });
		else assert.equal(await readFile(f.output, "utf8"), "tracked output");
	});
}

test("missing optional executable identifies docxquisite and PATH without output uncertainty", async (t) => {
	const f = await fixture(t, "success");
	await rm(f.executable);
	const error = await rejected(call(f.input, f.dir));
	assert.match(error.message, /docxquisite/i);
	assert.match(error.message, /PATH/);
	assert.match(error.message, /optional/i);
	assert.doesNotMatch(
		error.message,
		/(?:may|might|possibly|could)\s+(?:have been\s+)?(?:created|exist|written)/i,
	);
	await assert.rejects(readFile(f.output), { code: "ENOENT" });
});

test("closed stdin rejects a large request without leaking secrets", async (t) => {
	const f = await fixture(t, "closed-stdin");
	const largeSecret = secret + "x".repeat(900_000);
	const error = await rejected(
		call({ ...f.input, tracked_replace: { ...f.input.tracked_replace, old: largeSecret } }, f.dir),
	);
	assert.equal(await readFile(join(f.dir, "closed"), "utf8"), "yes");
	const exposed = `${error.stack} ${JSON.stringify(error, Object.getOwnPropertyNames(error))}`;
	for (const value of [secret, "PRIVATE replacement", "PRIVATE author"]) {
		assert.ok(!exposed.includes(value), "error must not expose request secrets");
	}
});
