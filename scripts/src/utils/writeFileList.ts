import { writeFileText } from "./writeFileText";

/**
 * write a list to a file
 * @param path - path to write the file to
 * @param list - string list to write to the file
 */
export function writeFileList(path: string, list: readonly string[]): void {
    const joined: string = list.join("\n");
    writeFileText(path, joined);
}
