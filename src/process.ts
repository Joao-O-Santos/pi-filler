import { spawn } from "node:child_process";

export const DEFAULT_COMMAND_TIMEOUT_MS = 30_000;

export interface CommandResult {
	stdout: Buffer;
	stderr: Buffer;
	code: number;
}

export class MissingExecutableError extends Error {
	constructor(public readonly executable: string) {
		super(
			`Required executable "${executable}" was not found. ` +
				`Install it and make sure it is available on PATH.`,
		);
		this.name = "MissingExecutableError";
	}
}

export interface RunCommandOptions {
	cwd?: string;
	env?: NodeJS.ProcessEnv;
	input?: string | Buffer;
	signal?: AbortSignal;
	timeoutMs?: number;
}

export function runCommand(
	executable: string,
	args: readonly string[],
	options: RunCommandOptions = {},
): Promise<CommandResult> {
	return new Promise((resolve, reject) => {
		if (options.signal?.aborted) {
			reject(new Error("Operation aborted"));
			return;
		}

		const child = spawn(executable, args, {
			cwd: options.cwd,
			env: options.env,
			stdio: ["pipe", "pipe", "pipe"],
		});
		const stdout: Buffer[] = [];
		const stderr: Buffer[] = [];
		let settled = false;
		let timedOut = false;

		const finish = (callback: () => void) => {
			if (settled) return;
			settled = true;
			if (timer) clearTimeout(timer);
			options.signal?.removeEventListener("abort", onAbort);
			callback();
		};
		const onAbort = () => {
			child.kill();
			finish(() => reject(new Error("Operation aborted")));
		};
		const timer = setTimeout(() => {
			timedOut = true;
			child.kill();
		}, options.timeoutMs ?? DEFAULT_COMMAND_TIMEOUT_MS);

		child.stdout.on("data", (chunk: Buffer) => stdout.push(chunk));
		child.stderr.on("data", (chunk: Buffer) => stderr.push(chunk));
		child.on("error", (error: NodeJS.ErrnoException) => {
			finish(() => {
				if (error.code === "ENOENT") reject(new MissingExecutableError(executable));
				else reject(new Error(`Failed to run ${executable}: ${error.message}`));
			});
		});
		child.on("close", (code) => {
			finish(() => {
				if (timedOut) {
					reject(
						new Error(
							`${executable} timed out after ${options.timeoutMs ?? DEFAULT_COMMAND_TIMEOUT_MS}ms`,
						),
					);
					return;
				}
				resolve({ stdout: Buffer.concat(stdout), stderr: Buffer.concat(stderr), code: code ?? -1 });
			});
		});

		options.signal?.addEventListener("abort", onAbort, { once: true });
		if (options.input !== undefined) child.stdin.end(options.input);
		else child.stdin.end();
	});
}

export function commandError(result: CommandResult, executable: string): Error {
	const message = result.stderr.toString("utf8").trim();
	return new Error(message || `${executable} exited with code ${result.code}`);
}
