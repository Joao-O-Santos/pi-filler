import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import { MissingExecutableError, runCommand } from "./process.js";

export interface TrackedReplacement {
	old: string;
	new: string;
	author: string;
	expected_sha256?: string;
}

const hashPattern = /^[0-9a-f]{64}$/;
const unknownOutcome =
	"docxquisite transport or protocol failure; output outcome may be unknown. " +
	"Inspect the output before any retry; no output was deleted.";
const abortedBeforeReplacement = "Operation aborted; no replacement was attempted.";
const abortedAfterReplacement =
	"Operation aborted; output outcome may be unknown. Inspect the output before any retry; no output was deleted.";

function checkAbort(signal: AbortSignal | undefined, mutationStarted: boolean): void {
	if (signal?.aborted) {
		throw new Error(mutationStarted ? abortedAfterReplacement : abortedBeforeReplacement);
	}
}

async function fingerprint(path: string, signal?: AbortSignal): Promise<string> {
	const hash = createHash("sha256");
	for await (const chunk of createReadStream(path, { signal })) hash.update(chunk);
	return hash.digest("hex");
}

const errorCodes = new Set([
	"invalid_request",
	"invalid_json",
	"request_too_large",
	"unsupported_version",
	"unknown_operation",
	"not_found",
	"output_exists",
	"permission_denied",
	"document_error",
	"io_error",
	"internal_error",
]);

function object(value: unknown): value is Record<string, unknown> {
	return typeof value === "object" && value !== null && !Array.isArray(value);
}

function exactKeys(value: Record<string, unknown>, keys: string[]): boolean {
	return Object.keys(value).length === keys.length && keys.every((key) => key in value);
}

function strings(value: unknown): value is string[] {
	return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

export function validateTrackedReplacement(value: unknown): asserts value is TrackedReplacement {
	if (
		!object(value) ||
		Object.keys(value).some((key) => !["old", "new", "author", "expected_sha256"].includes(key)) ||
		typeof value.old !== "string" ||
		value.old.length === 0 ||
		typeof value.new !== "string" ||
		typeof value.author !== "string" ||
		value.author.length === 0 ||
		(value.expected_sha256 !== undefined &&
			(typeof value.expected_sha256 !== "string" || !hashPattern.test(value.expected_sha256)))
	) {
		throw new Error(
			"tracked_replace requires nonempty old and author, string new, and an optional lowercase SHA-256",
		);
	}
}

async function protocol(
	request: Record<string, unknown>,
	cwd: string,
	signal?: AbortSignal,
): Promise<Record<string, unknown>> {
	let response: unknown;
	let code: number;
	const probing = request.operation === "capabilities";
	const failure = probing
		? "docxquisite capability probe failed; no replacement was attempted."
		: unknownOutcome;
	checkAbort(signal, !probing);
	try {
		const stdin = JSON.stringify({ version: 1, ...request });
		if (Buffer.byteLength(stdin, "utf8") > 1024 * 1024) {
			throw new Error("Request too large");
		}
		const result = await runCommand("docxquisite", ["protocol"], {
			cwd,
			signal,
			stdin,
			timeoutMs: 30_000,
			maxBufferBytes: 64 * 1024,
		});
		code = result.code;
		response = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(result.stdout));
	} catch (error) {
		checkAbort(signal, !probing);
		// Neither command errors, stderr, nor returned diagnostic text are safe to forward.
		if (probing && error instanceof MissingExecutableError) {
			throw new Error(
				"Optional executable docxquisite not found; install it and make sure it is available on PATH.",
			);
		}
		throw new Error(failure);
	}
	checkAbort(signal, !probing);
	if (!object(response) || response.version !== 1 || typeof response.ok !== "boolean") {
		throw new Error(failure);
	}
	if (response.ok) {
		if (
			code !== 0 ||
			!exactKeys(response, ["version", "ok", "result"]) ||
			!object(response.result)
		) {
			throw new Error(failure);
		}
		return response.result;
	}
	if (
		code !== 1 ||
		!exactKeys(response, ["version", "ok", "error"]) ||
		!object(response.error) ||
		!exactKeys(response.error, ["code", "message"]) ||
		typeof response.error.code !== "string" ||
		!errorCodes.has(response.error.code) ||
		typeof response.error.message !== "string"
	) {
		throw new Error(failure);
	}
	throw new Error(
		`docxquisite rejected the operation (${response.error.code}). ` +
			(probing
				? "No replacement was attempted."
				: "No retry or cleanup was performed; inspect any output before retrying."),
	);
}

export async function replaceTracked(
	input: string,
	output: string,
	replacement: TrackedReplacement,
	cwd: string,
	signal?: AbortSignal,
): Promise<{ output: string; source_sha256: string; output_sha256: string }> {
	checkAbort(signal, false);
	validateTrackedReplacement(replacement);
	if (input === output) throw new Error("DOCX tracked replacement output must differ from input");
	const capabilities = await protocol({ operation: "capabilities" }, cwd, signal);
	if (
		!exactKeys(capabilities, ["protocol_version", "operations", "replace_limitations"]) ||
		capabilities.protocol_version !== 1 ||
		!strings(capabilities.operations) ||
		!capabilities.operations.includes("replace_tracked") ||
		!strings(capabilities.replace_limitations)
	) {
		throw new Error(
			"docxquisite does not advertise the required protocol v1 tracked replacement capability",
		);
	}
	let sourceHash: string;
	try {
		sourceHash = await fingerprint(input, signal);
	} catch {
		checkAbort(signal, false);
		throw new Error("Unable to fingerprint tracked replacement input");
	}
	checkAbort(signal, false);
	if (replacement.expected_sha256 !== undefined && replacement.expected_sha256 !== sourceHash) {
		throw new Error("Tracked replacement source SHA-256 mismatch");
	}
	checkAbort(signal, false);
	const result = await protocol(
		{ operation: "replace_tracked", input, output, ...replacement, expected_sha256: sourceHash },
		cwd,
		signal,
	);
	if (
		!exactKeys(result, ["output", "source_sha256", "output_sha256", "checks", "limitations"]) ||
		result.output !== output ||
		result.source_sha256 !== sourceHash ||
		typeof result.output_sha256 !== "string" ||
		!hashPattern.test(result.output_sha256) ||
		!strings(result.checks) ||
		!strings(result.limitations)
	) {
		throw new Error(unknownOutcome);
	}
	try {
		const outputPath = resolve(cwd, output);
		if (!(await stat(outputPath)).isFile()) throw new Error("Not a regular file");
		const outputHash = await fingerprint(outputPath, signal);
		checkAbort(signal, true);
		if (outputHash !== result.output_sha256) throw new Error("Output hash mismatch");
	} catch {
		checkAbort(signal, true);
		throw new Error(unknownOutcome);
	}
	checkAbort(signal, true);
	// Return only locally known paths and validated fingerprints, never toolkit prose.
	return { output, source_sha256: sourceHash, output_sha256: result.output_sha256 };
}
