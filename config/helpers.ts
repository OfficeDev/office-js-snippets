import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as chalk from 'chalk';
import { Observable } from 'rxjs/Observable';
import 'rxjs/add/operator/mergeMap';
import 'rxjs/add/observable/from';
import 'rxjs/add/observable/of';
import { kebabCase } from 'lodash';

export interface File {
    name: string;
    path: string;
    source: string;
    group: string;
}

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
 * @param file An absolute path to the file.
  * @param root An absolute path to the root directory.
 */
export const processFile = (file: string, root: string) => {
    /* Determine the platform as windows uses '\' where as linux uses '/' */
    const delimiter = os.platform() === 'win32' ? '\\' : '/';

    /* Get the relative path to the file from the root directory '/' */
    const relativePath = path.relative(root, file);

    /* Extract the required properties */
    let [name, group, source, ...additional] = relativePath.split(delimiter).reverse();

    /* Additional must be null or empty */
    if (additional && additional.length > 0) {
        throw new Error(`Invalid folder structure at ${chalk.bold.red(relativePath)}. File ${chalk.bold.yellow(name)} was located too deep.`);
    }

    if (source == null) {
        source = group;
    }

    return Observable.of<File>({
        path: relativePath,
        source,
        group,
        name
    });
};

/**
 * Recurrsively crawl through a folder and return all the files in it.
 * @param dir An absolute path to the directory.
 * @param root An absolute path to the root directory.
 */
export const getFiles = (dir: string, root: string): Observable<File> =>
    /*
    * Convert all the files into an Observable stream of files.
    * This allows us to focus the remainder of the operations
    * on a PER FILE basis.
    */
    Observable
        .from(readDir(dir))
        .mergeMap(files => Observable.from(files))
        .mergeMap((file) => {
            const filePath = path.join(dir, file);
            const withoutExt = file.replace('.yaml', '');

            /* Check for file/folder naming guidelines */
            if (kebabCase(withoutExt) !== withoutExt) {
                throw new Error(`Invalid name at ${chalk.bold.red(filePath)}. Name was expected to be ${chalk.bold.magenta(kebabCase(withoutExt))}, found ${chalk.bold.yellow(withoutExt)}.`);
            }

            /*
            * Check if the file is a folder and either return
            * an Observable to the recurrsive walk operation or
            * return an Observable of the file object itself.
            */
            return Observable
                .from(isDir(filePath))
                .mergeMap(pathIsDirectory =>
                    pathIsDirectory ?
                        getFiles(filePath, root) :
                        processFile(filePath, root)
                );
        });

