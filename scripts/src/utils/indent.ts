import { lineSplit } from "./lineSplit";
import { mergeWithDefaults } from "./mergeWithDefaults";

/**
 * describe a single level of indent
 */
export interface IndentOptions {
    /**
     * the value to use for the indent
     * default of four spaces
     */
    value: string;

    /**
     * the number of the value to use for a single level of indent
     * default of 1
     */
    count: number;

    /**
     * the number of times to indent
     * default of 1
     */
    level: number;
}

const defaultIndent: IndentOptions = {
    value: " ",
    count: 4,
    level: 1,
};

/**
 * indent all lines with the specified level of indent.
 * @param string - string to indent
 * @param indent - indent options
 * @returns a version of the string indented according to the indent options
 */
export function indent(string: string, indent: Partial<IndentOptions> = defaultIndent): string {
    const settings: IndentOptions = mergeWithDefaults(indent, defaultIndent);

    const indentString = settings.value.repeat(settings.count).repeat(settings.level);

    // this also indents any empty lines
    return indentString + lineSplit(string).join(`\n${indentString}`);
}
