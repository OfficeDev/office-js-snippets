/**
 * decapitalize a string
 * @param string - string to decapitalize
 */
export function decapitalize(string: string): string {
    if (string.length === 0) {
        return string;
    }
    return string.charAt(0).toLowerCase() + string.slice(1);
}
