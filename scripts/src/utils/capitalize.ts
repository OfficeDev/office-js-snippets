/**
 * uppercases the first character in a string.
 * in the case that the first character in the string can not be upper cased (for example a white space character or an empty string) the string is unmodified.
 * @param word - string to capitalize
 * @returns the string with it's first character upper cased.
 *
 */
export function capitalize(word: string): string {
    if (!word || word.length === 0) {
        return word;
    }

    if (word.length === 1) {
        return word.toUpperCase();
    }

    return word.substring(0, 1).toUpperCase() + word.substring(1);
}
