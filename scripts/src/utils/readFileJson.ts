import { readFileText } from "./readFileText";

/**
 * Read a file that contains JSON and turn it into an object
 *
 * Note: no validation is done on the data.
 *
 * @param path - path to the JSON file
 */
export function readFileJson<T>(path: string): T {
    const data: string = readFileText(path);
    const object: T = JSON.parse(data);
    return object;
}
