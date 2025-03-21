import { standardizeNewlines } from "./standardizeNewlines";

/**
 * split a string into a list of lines
 * @param string - string to split
 * @returns list of the individual lines in the string
 */
export function lineSplit(string: string): string[] {
    return standardizeNewlines(string).split("\n");
}
