import { execFile } from "node:child_process";

export const DEFAULT_COMMAND_TIMEOUT_MS = 30_000;
export const DEFAULT_COMMAND_MAX_BUFFER_BYTES = 10 * 1024 * 1024;

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
	signal?: AbortSignal;
	timeoutMs?: number;
	maxBufferBytes?: number;
}

type ExecFailure = Error & {
	code?: string | number | null;
	killed?: boolean;
};

export function runCommand(
	executable: string,
	args: readonly string[],
	options: RunCommandOptions = {},
): Promise<CommandResult> {
	return new Promise((resolve, reject) => {
		execFile(
			executable,
			[...args],
			{
				cwd: options.cwd,
				env: options.env,
				signal: options.signal,
				timeout: options.timeoutMs ?? DEFAULT_COMMAND_TIMEOUT_MS,
				maxBuffer: options.maxBufferBytes ?? DEFAULT_COMMAND_MAX_BUFFER_BYTES,
				killSignal: "SIGTERM",
				encoding: null,
			},
			(error, stdout, stderr) => {
				if (!error) {
					resolve({ stdout, stderr, code: 0 });
					return;
				}

				const failure = error as ExecFailure;
				if (failure.code === "ENOENT") {
					reject(new MissingExecutableError(executable));
					return;
				}
				if (failure.code === "ABORT_ERR") {
					reject(new Error("Operation aborted"));
					return;
				}
				if (failure.code === "ERR_CHILD_PROCESS_STDIO_MAXBUFFER") {
					reject(
						new Error(
							`${executable} output exceeded ${options.maxBufferBytes ?? DEFAULT_COMMAND_MAX_BUFFER_BYTES} bytes`,
						),
					);
					return;
				}
				if (failure.killed) {
					reject(
						new Error(
							`${executable} timed out after ${options.timeoutMs ?? DEFAULT_COMMAND_TIMEOUT_MS}ms`,
						),
					);
					return;
				}
				if (typeof failure.code === "number") {
					resolve({ stdout, stderr, code: failure.code });
					return;
				}
				reject(new Error(`Failed to run ${executable}: ${failure.message}`));
			},
		);
	});
}

export function commandError(result: CommandResult, executable: string): Error {
	const message = result.stderr.toString("utf8").trim();
	return new Error(message || `${executable} exited with code ${result.code}`);
}
