/**
 * standardize newlines to proper unix line endings
 * @param string - string to standardize
 */
export function standardizeNewlines(string: string): string {
    return string.replace(/\r/gm, "");
}
