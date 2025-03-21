/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * create a new object that ensure all default properties are present
 * @param original - original object
 * @param defaults - default object
 */
export function mergeWithDefaults<T extends object>(original: Partial<T>, defaults: T): T {
    const o: any = original;
    const d: any = defaults;
    const merge: any = {}; //shallowCopyOwnProperties(original);

    Object.getOwnPropertyNames(defaults).forEach((name) => {
        merge[name] = Object.getOwnPropertyDescriptor(o, name) ? o[name] : d[name];
    });

    return merge;
}
