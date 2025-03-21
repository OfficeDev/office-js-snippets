import { lineSplit } from "./lineSplit";
import { listWithoutDuplicateElements } from "./listWithoutDuplicateElements";
import { readFileText } from "./readFileText";

/**
 * reads lines from a file and removes the ones that are whitespace.
 * @param path - path to read the file from
 */
export function readFileList(path: string): string[] {
    const data: string = readFileText(path);
    return listWithoutDuplicateElements(lineSplit(data));
}
