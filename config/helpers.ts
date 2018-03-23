import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as chalk from 'chalk';
import * as jsyaml from 'js-yaml';
import { console } from './status';
import * as rimraf from 'rimraf';

export const officeHostsToAppNames = {
    'ACCESS': 'Access',
    'EXCEL': 'Excel',
    'ONENOTE': 'OneNote',
    'OUTLOOK': 'Outlook',
    'POWERPOINT': 'PowerPoint',
    'PROJECT': 'Project',
    'WORD': 'Word'
};

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
    endpoints?: string[];
}

/**
 * Creates a chalk based section with the specific color.
 * @param title Title of the banner.
 * @param message Message of the banner.
 * @param chalkFunction Chalk color function.
 */
export const banner = (title: string, message: string = null, chalkFn: chalk.ChalkChain = null) => {
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
    const details = jsyaml.safeDump(item, {
        indent: 4,
        lineWidth: -1,
        skipInvalid: true
    });

    return details.split('\n').map(line => new Array(indent).join(' ') + line).join('\n');
}

export const getDestinationBranch = (sourceBranch: 'master' | 'prod' | any): 'deploy-beta' | 'deploy-prod' | null => {
    if (sourceBranch === 'master') {
        return 'deploy-beta';
    }
    else if (sourceBranch === 'prod') {
        return 'deploy-prod';
    }
    else {
        return null;
    }
};

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
export const rmRf = (dir: string) =>
    new Promise<string>((resolve, reject) => {
        const location = path.resolve(dir);
        rimraf(location, (err) => {
            if (err) {
                return reject(err);
            }
            return resolve(location);
        });
    });

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
    new Promise((resolve, reject) => {
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
 * Load the contents of the YAML file.
 * @param path Absolute to the yaml file.
 */
export const loadFileContents = (path: string) =>
    new Promise<string>(async (resolve, reject) => {
        let pathIsDirectory = await isDir(path);
        if (pathIsDirectory) {
            return reject(new Error(`Cannot open a directory @ ${chalk.bold.red(path)}`));
        }
        else {
            fs.readFile(path, 'UTF8', (err, contents) => {
                try {
                    if (err) {
                        return reject(err);
                    }
                    return resolve(contents);
                }
                catch (err) {
                    reject(err);
                }
            });
        }
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
        throw new Error(`Invalid folder structure at ${chalk.bold.red(relativePath)}.File ${chalk.bold.yellow(name)} was located too deep.`);
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
 * Recurrsively crawl through a folder and return all the files in it.
 * @param root An absolute path to the directory.
 */
export function getFiles(root: string): SnippetFileInput[] {
    let files: SnippetFileInput[] = [];
    syncRecurseThroughDirectory(root);
    return files;


    // Helper
    function syncRecurseThroughDirectory(dir: string) {
        fs.readdirSync(dir)
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
