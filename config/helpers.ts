import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import chalk from 'chalk';
import * as jsyaml from 'js-yaml';
import { rimraf } from 'rimraf';
import { isObject, isNil, isString, isEmpty } from 'lodash';

export interface SnippetFileInput {
    file_name: string;
    relativePath: string;
    fullPath: string;
    host: string;
    group: string;
    isPublic: boolean;
}

export interface SnippetProcessedData {
    id: string;
    name: string;
    fileName: string;
    relativePath: string;
    fullPath: string;
    description: string;
    host: string;
    rawUrl: string;
    group: string;
    order: number;
    api_set: {
        [index: string]: number
    };
    isPublic: boolean;
}

/**
 * Creates a chalk based section with the specific color.
 * @param title Title of the banner.
 * @param message Message of the banner.
 * @param chalkFunction Chalk color function.
 */
export const banner = (title: string, message: string = null, chalkFn: any = null) => {
    if (!chalkFn) {
        chalkFn = chalk.bold;
    }

    const dashes = Array(Math.max(title.length + 1, 30)).join('-');
    console.log(chalkFn(`\n\n${dashes}`));
    console.log(chalkFn(`${title}`));
    if (message) {
        console.log(chalkFn(dashes));
        console.log(message);
    }
    console.log(chalkFn(`${dashes}\n`));
};

export function getPrintableDetails(item: any, indent: number) {
    const details = jsyaml.dump(item, {
        indent: 4,
        lineWidth: -1,
        skipInvalid: true
    });

    return details.split('\n').map(line => new Array(indent).join(' ') + line).join('\n');
}

/**
 * Creates a folder.
 * @param dir An absolute path to the directory.
 */
export const mkDir = (dir: string) =>
    new Promise<string>((resolve, reject) => {
        const location = path.resolve(dir);
        fs.mkdir(location, (err) => {
            if (err) {
                return reject(err);
            }
            return resolve(location);
        });
    });

/**
* Deletes a folder.
* @param dir An absolute path to the directory.
*/
export const rmRf = async (dir: string): Promise<string> => {
    const location = path.resolve(dir);
    await rimraf(location);
    return location;
};

/**
 * Load all the files and folders in a given directory.
 * @param dir An absolute path to the directory.
 */
export const readDir = (dir: string) =>
    new Promise<string[]>((resolve, reject) => {
        fs.readdir(dir, (err, files) => {
            if (err) {
                return reject(err);
            }
            return resolve(files);
        });
    });

/**
 * Write to file
 * @param filename
 * @param contents
 */
export const writeFile = (filename: string, contents: string) =>
    new Promise<void>((resolve, reject) => {
        fs.writeFile(filename, contents, (err) => {
            if (err) {
                return reject(err);
            }
            return resolve();
        });
    });

/**
 * Check whether the given path is a file or a directory.
 * @param path An absolute path to the directory.
 */
export const isDir = (path: string) =>
    new Promise<boolean>((resolve, reject) => {
        fs.stat(path, (err, file) => {
            if (err) {
                return reject(err);
            }
            return resolve(file.isDirectory());
        });
    });

/**
 * Check the file path against validations and return a 'File' object.
 * @param fullPath An absolute path to the file.
  * @param root An absolute path to the root directory.
 */
export function getFileMetadata(fullPath: string, root: string): SnippetFileInput {
    /* Determine the platform as windows uses '\' where as linux uses '/' */
    const delimiter = os.platform() === 'win32' ? '\\' : '/';

    /* Get the relative path to the file from the root directory '/' */
    const relativePath = path.relative(root, fullPath);

    /* Extract the required properties */
    let [file_name, group, host, ...additional] = relativePath.split(delimiter).reverse();

    /* Additional must be null or empty */
    if (additional && additional.length > 0) {
        throw new Error(`Invalid folder structure at ${chalk.bold.red(relativePath)}.File ${chalk.bold.yellow(file_name)} was located too deep.`);
    }

    if (host == null) {
        host = group;
    }

    host = host.toLowerCase();

    return {
        relativePath: relativePath,
        fullPath,
        isPublic: !(/[\\/]private-samples$/.test(root)),
        host,
        group,
        file_name
    };
}

/**
 * Recursively crawl through a folder and return all the files in it.
 * @param root An absolute path to the directory.
 */
export function getFiles(root: string): SnippetFileInput[] {
    let files: SnippetFileInput[] = [];
    syncRecurseThroughDirectory(root);
    return files;


    // Helper
    function syncRecurseThroughDirectory(dir: string) {
        fs.readdirSync(dir)
        .filter(file => !['.DS_Store'].includes(file))
        .forEach(file => {
            const fullPath = path.join(dir, file);
            const withoutExt = file.replace('.yaml', '');

            /* Check for file/folder naming guidelines */
            if (!followsNamingGuidelines(withoutExt)) {
                throw new Error(`Invalid name at ${chalk.bold.red(fullPath)}. Name must only contain lowercase letters, numbers, and hyphens.`);
            }

            if (fs.statSync(fullPath).isDirectory()) {
                syncRecurseThroughDirectory(fullPath);
            } else {
                files.push(getFileMetadata(fullPath, root));
            }
        });
    }
}

/**
    Naming guidelines:  only allow lowercase letters, numbers, and hyphens

    OK:
    sample
    sample-with-hyphen
    sample-es5


    BAD:
    sample with space
    Any-uppercase
    anyWhere
    or_underscores
    or.dots
    $likethistoo
*/
export function followsNamingGuidelines(name: string) {
    return /^[a-z0-9\-]+$/.test(name);
}

/** Determines if a name is really just a 25-character machine-generated ID */
export function isCUID(id: string) {
    if (id.trim().length === 25 && id.indexOf('_') === -1) {
        // not likely to be a real id, with a name of that precise length and all as one word.
        return true;
    }

    return false;
}

export function pause(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Helper for creating and querying Dictionaries.
 * A wrapper around ES6 Maps.
 */
export class Dictionary<T> {
    protected _items: Map<string, T>;

    /**
     * @constructor
     * @param {object} items Initial seed of items.
     */
    constructor(items?: { [index: string]: T } | Array<[string, T]> | Map<string, T>) {
        if (isNil(items)) {
            this._items = new Map();
        }
        else if (items instanceof Set) {
            throw new TypeError(`Invalid type of argument: Set`);
        }
        else if (items instanceof Map) {
            this._items = new Map(items);
        }
        else if (Array.isArray(items)) {
            this._items = new Map(items);
        }
        else if (isObject(items)) {
            this._items = new Map();
            for (const key of Object.keys(items)) {
                this._items.set(key, items[key]);
            }
        }
        else {
            throw new TypeError(`Invalid type of argument: ${typeof items}`);
        }
    }

    /**
     * Gets an item from the dictionary.
     *
     * @param {string} key The key of the item.
     * @return {object} Returns an item if found.
     */
    get(key: string): T {
        return this._items.get(key);
    }

    /**
     * Inserts an item into the dictionary.
     * If an item already exists with the same key, it will be overridden by the new value.
     *
     * @param {string} key The key of the item.
     * @param {object} value The item to be added.
     * @return {object} Returns the added item.
     */
    set(key: string, value: T): T {
        this._validateKey(key);
        this._items.set(key, value);
        return value;
    }

    /**
     * Removes an item from the dictionary.
     * Will throw if the key doesn't exist.
     *
     * @param {string} key The key of the item.
     * @return {object} Returns the deleted item.
     */
    delete(key: string): T {
        if (!this.has(key)) {
            throw new ReferenceError(`Key: ${key} not found.`);
        }
        let value = this._items.get(key);
        this._items.delete(key);
        return value;
    }

    /**
     * Clears the dictionary.
     */
    clear(): void {
        this._items.clear();
    }

    /**
     * Check if the dictionary contains the given key.
     *
     * @param {string} key The key of the item.
     * @return {boolean} Returns true if the key was found.
     */
    has(key: string): boolean {
        this._validateKey(key);
        return this._items.has(key);
    }

    /**
     * Lists all the keys in the dictionary.
     *
     * @return {array} Returns all the keys.
     */
    keys(): Array<string> {
        return Array.from(this._items.keys());
    }

    /**
     * Lists all the values in the dictionary.
     *
     * @return {array} Returns all the values.
     */
    values(): Array<T> {
        return Array.from(this._items.values());
    }

    /**
     * Get a shallow copy of the underlying map.
     *
     * @return {object} Returns the shallow copy of the map.
     */
    clone(): Map<string, T> {
        return new Map(this._items);
    }

    /**
     * Number of items in the dictionary.
     *
     * @return {number} Returns the number of items in the dictionary.
     */
    get count(): number {
        return this._items.size;
    }

    private _validateKey(key: string): void {
        if (!isString(key) || isEmpty(key)) {
            throw new TypeError('Key needs to be a string');
        }
    }
}
