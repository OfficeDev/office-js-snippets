import { capitalize } from "./capitalize";

/**
 * splits on whitespaces and -, capitalizes words, and joins them
 * @param words
 */
export function joinWords(words: string): string {
    return words
        .split(/(\s|-)/)
        .map((word) => capitalize(word))
        .join("");
}
