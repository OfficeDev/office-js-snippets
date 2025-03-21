/**
 * creates list without duplicates from an original list, comparing using the default comparison operator, keeping only the first occurrence.
 * @param original - list to
 * @returns new list without the duplicates present in the original
 */
export function listWithoutDuplicateElements<T>(original: readonly T[]): T[] {
    // only take the first item
    return original.filter(
        (value: T, index: number, array: readonly T[]) => array.indexOf(value) === index,
    );
}
