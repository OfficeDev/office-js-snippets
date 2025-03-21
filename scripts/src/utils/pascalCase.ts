import { joinWords } from "./joinWords";
import { capitalize } from "./capitalize";

/**
 * PascalCase
 * splits on spaces and capitalizes words in between
 * @param string - string to pascalCase
 */
export function pascalCase(string: string): string {
    string = joinWords(string);
    return capitalize(string);
}
