import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test, { type TestContext } from "node:test";
import { executeFiller, type FillerInput } from "../src/index.js";

const secrets = ["PRIVATE old 7c9", "PRIVATE new 8d0", "PRIVATE author 9e1"];

async function fixture(t: TestContext, hold: "capabilities" | "replace_tracked") {
	const dir = await mkdtemp(join(tmpdir(), "filler-cancel-"));
	const previousPath = process.env.PATH;
	const previousDir = process.env.CANCEL_TEST_DIR;
	const previousHold = process.env.CANCEL_TEST_HOLD;
	t.after(async () => {
		try {
			const pid = Number(await readFile(join(dir, "pid"), "utf8"));
			try {
				process.kill(pid, "SIGKILL");
			} catch (error) {
				if ((error as NodeJS.ErrnoException).code !== "ESRCH") throw error;
			}
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		}
		if (previousPath === undefined) delete process.env.PATH;
		else process.env.PATH = previousPath;
		if (previousDir === undefined) delete process.env.CANCEL_TEST_DIR;
		else process.env.CANCEL_TEST_DIR = previousDir;
		if (previousHold === undefined) delete process.env.CANCEL_TEST_HOLD;
		else process.env.CANCEL_TEST_HOLD = previousHold;
		await rm(dir, { recursive: true, force: true });
	});
	process.env.PATH = dir;
	process.env.CANCEL_TEST_DIR = dir;
	process.env.CANCEL_TEST_HOLD = hold;
	await writeFile(
		join(dir, "docxquisite"),
		`#!${process.execPath}
const fs = require('node:fs');
const path = require('node:path');
const dir = process.env.CANCEL_TEST_DIR;
let raw = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => raw += chunk);
process.stdin.on('end', () => {
  const request = JSON.parse(raw);
  fs.appendFileSync(path.join(dir, 'calls'), request.operation + '\\n');
  if (request.operation === process.env.CANCEL_TEST_HOLD) {
    if (request.operation === 'replace_tracked') {
      fs.writeFileSync(request.output, 'partial output');
      console.error(request.old + request.new + request.author);
    }
    fs.writeFileSync(path.join(dir, 'pid'), String(process.pid));
    fs.writeFileSync(path.join(dir, 'ready'), request.operation);
    setInterval(() => {}, 1000);
    return;
  }
  console.log(JSON.stringify({version:1,ok:true,result:{protocol_version:1,operations:['capabilities','replace_tracked'],replace_limitations:[]}}));
});
`,
	);
	await chmod(join(dir, "docxquisite"), 0o755);
	await writeFile(join(dir, "source.docx"), "source bytes");
	const input = {
		format: "docx",
		action: "patch",
		path: "@source.docx",
		output: "@output.docx",
		tracked_replace: { old: secrets[0], new: secrets[1], author: secrets[2] },
	} as FillerInput;
	return { dir, input, output: join(dir, "output.docx") };
}

async function waitReady(dir: string) {
	const deadline = Date.now() + 5000;
	while (Date.now() < deadline) {
		try {
			return await readFile(join(dir, "ready"), "utf8");
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		}
		await new Promise((resolve) => setTimeout(resolve, 10));
	}
	throw new Error("fake executable did not reach handshake");
}

async function calls(dir: string) {
	try {
		return (await readFile(join(dir, "calls"), "utf8")).trim().split("\n");
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
		throw error;
	}
}

async function cancellation(promise: Promise<unknown>) {
	const result = await promise.then(
		() => undefined,
		(error: unknown) => error,
	);
	assert.ok(result instanceof Error, "operation must reject");
	assert.match(result.message, /abort|cancel/i);
	const exposed = `${result.stack} ${JSON.stringify(result, Object.getOwnPropertyNames(result))}`;
	for (const secret of secrets) assert.ok(!exposed.includes(secret), "error must not leak secrets");
	return result;
}

test("pre-aborted tracked replacement never attempts mutation", async (t) => {
	const f = await fixture(t, "capabilities");
	const controller = new AbortController();
	controller.abort();
	await cancellation(executeFiller(f.input, f.dir, controller.signal));
	assert.ok(!(await calls(f.dir)).includes("replace_tracked"));
	await assert.rejects(readFile(f.output), { code: "ENOENT" });
});

test("cancellation during capabilities does not attempt mutation", async (t) => {
	const f = await fixture(t, "capabilities");
	const controller = new AbortController();
	const result = cancellation(executeFiller(f.input, f.dir, controller.signal));
	void result.catch(() => {});
	try {
		assert.equal(await waitReady(f.dir), "capabilities");
	} finally {
		controller.abort();
	}
	await result;
	assert.deepEqual(await calls(f.dir), ["capabilities"]);
	await assert.rejects(readFile(f.output), { code: "ENOENT" });
});

test("cancellation during mutation warns of unknown output and preserves it", async (t) => {
	const f = await fixture(t, "replace_tracked");
	const controller = new AbortController();
	const result = cancellation(executeFiller(f.input, f.dir, controller.signal));
	void result.catch(() => {});
	try {
		assert.equal(await waitReady(f.dir), "replace_tracked");
	} finally {
		controller.abort();
	}
	const error = await result;
	assert.match(
		error.message,
		/(?:output.*(?:may|might|could|unknown)|(?:may|might|could|unknown).*output)/i,
	);
	assert.deepEqual(await calls(f.dir), ["capabilities", "replace_tracked"]);
	assert.equal(await readFile(f.output, "utf8"), "partial output");
	assert.equal(await readFile(join(f.dir, "source.docx"), "utf8"), "source bytes");
});
