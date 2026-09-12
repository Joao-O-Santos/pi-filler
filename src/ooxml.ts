import { unzipSync, zipSync } from "fflate";

export type OfficePackage = Record<string, Uint8Array>;

export function unpackOfficePackage(bytes: Uint8Array, format: string): OfficePackage {
	try {
		return unzipSync(bytes);
	} catch {
		throw new Error(`Invalid ${format} ZIP package`);
	}
}

export function packOfficePackage(pkg: OfficePackage): Uint8Array {
	return zipSync(pkg);
}
