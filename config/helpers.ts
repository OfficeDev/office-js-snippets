import * as path from 'path';
import * as fs from 'fs';

const readdir = (dir: string) =>
    new Promise<string[]>((resolve, reject) => {
        fs.readdir(dir, (err, files) => {
            if (err) {
                return reject(err);
            }
            return resolve(files);
        });
    });

const isDir = (path: string) =>
    new Promise<boolean>((resolve, reject) => {
        fs.stat(path, (err, file) => {
            if (err) {
                return reject(err);
            }
            return resolve(file.isDirectory());
        });
    });

const walk = async (dir: string): Promise<string[]> => {
    let currentFolder: string[] = [];
    const files = await readdir(dir);

    const promises = files.map(async (file) => {
        const pathIsDirectory = await isDir(path.join(dir, file));
        if (pathIsDirectory) {
            walk(path.join(dir, file));
        }
        else {
            currentFolder.push(path.join(dir, file));
        }
    });

    return await promises;
};
