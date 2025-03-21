/**
 * checks if two lists have the same values in the same order using the default comparison operator.
 *
 * @param a - a list
 * @param b - a list
 * @returns true if both lists have the same values in the same order.
 */
export function equivalentLists(a: string[], b: string[]): boolean {
    if (a.length !== b.length) {
        return false;
    }

    for (let i = 0; i < a.length; i++) {
        if (a[i] !== b[i]) {
            return false;
        }
    }

    return true;
}
