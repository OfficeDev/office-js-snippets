#!/usr/bin/env node --harmony

import * as path from 'path';
import * as chalk from 'chalk';
import { getFiles, File } from './helpers';

console.log(chalk.bold.yellow('Loading samples...'));
let files$ = getFiles(path.resolve('samples'), path.resolve('samples'));
files$.subscribe(
    (file: File) => console.log(file.name, file.group, file.source, file.path),
    (error: Error) => {
        console.log(chalk.bold.red('\n\n--------------------------------------'));
        console.log(chalk.bold.red('\tAn error has occurred'));
        console.log(chalk.bold.red('--------------------------------------\n\n'));
        console.log(error.message);
        console.log(chalk.bold.red('\n\n--------------------------------------\n\n'));
        process.exitCode = 1;
        throw error;
    }
);
// files = files.filter(file => !/default.yaml?$/i.test(file));
// let processedFiles = files.map(process);
// let groupedFiles = _.groupBy(processedFiles, 'source');

// fs.mkdirSync(path.resolve('playlists'));

// _.each(groupedFiles, (items, group) => {
//     let pluckedItems = _.map(items, 'meta');
//     let contents = jsYaml.safeDump(pluckedItems);
//     fs.writeFileSync(path.resolve(`playlists/${group}.yaml`), contents);
//     console.log(chalk.bold.yellow(`\nCreated ${group}.yaml`));
// });
