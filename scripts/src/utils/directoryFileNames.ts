import { readdirSync } from "fs";
import { isFile } from "./isFile";
import { join } from "path";

/**
 * retrieves the file names present in the directory
 * @param path - path of the directory to get the files in
 * @returns list of file names in the directory
 */
export function directoryFileNames(path: string): string[] {
    const all = readdirSync(path);
    const files = all.filter((file: string) => isFile(join(path, file)));
    // paths are sorted because determinism is convenient for testing and reproduction of issues.
    return files.sort();
}
