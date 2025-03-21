import { lstatSync } from "fs";

/**
 * is the path a file?
 *
 * @param path - path to test
 * @returns true when the path is a file
 */
export function isFile(path: string): boolean {
    return lstatSync(path).isFile();
}
