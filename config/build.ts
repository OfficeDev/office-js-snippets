#!/usr/bin/env node --harmony

import * as path from 'path';
import * as chalk from 'chalk';
import { status } from './status';
import { rmRf, mkDir, getFiles, writeFile, File, banner, loadYamlFile } from './helpers';
import { startCase, groupBy, map } from 'lodash';
import * as jsyaml from 'js-yaml';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/filter';

const { GH_ACCOUNT, GH_REPO } = process.env;
const files: File[] = [];

(async () => {
    try {
        /* Creating playlists directory */
        status.add('Creating \'playlists\' folder');
        await rmRf('playlists');
        await mkDir('playlists');
        status.complete('Creating \'playlists\' folder');

        /* Loading samples */
        status.add('Loading snippets');
        let files$ = getFiles(path.resolve('samples'), path.resolve('samples'));
        status.complete('Loading snippets');

        files$.mergeMap(async (file) => {
            try {
                status.add(`Processing ${file.host}::${file.name}`);
                let { name, description } = await loadYamlFile<{ name: string, description: string }>(path.resolve('samples', file.path));
                status.complete(`Processing ${file.host}::${file.name}`);
                return {
                    name,
                    description,
                    ...file,
                    gist: `https://raw.githubusercontent.com/${GH_ACCOUNT}/${GH_REPO}/deployment/samples/${file.host}/${file.group}/${file.name}`,
                    group: startCase(file.group)
                };
            } catch (exception) {
                status.complete(`Processing ${file.host}::${file.name}`, exception);
                handleError(`Failed to process ${file.host}::${file.name}: ${exception.message || exception}`);
                return null;
            }
        })
            .filter((file) => !(file == null) && file.name !== 'default.yaml')
            .map((file) => files.push(file))
            .subscribe(null, handleError, snippetsProcessed);
    }
    catch (exception) {
        handleError(exception);
    }
})();

/**
 * Generic error handler.
 * @param error Error object.
 */
function handleError(error?: any) {
    banner('An error has occured', error.message || error, chalk.bold.red);
    process.exit(1);
}

/**
 * Generating playlists
 */
async function snippetsProcessed() {
    if (files == null || files.length < 1) {
        return;
    }

    /* Generating playlists */
    status.add('Generating playlists');
    const groups = groupBy(files, 'host');
    let promises = map(groups, async (items, group) => {
        let contents = jsyaml.safeDump(items);
        await writeFile(path.resolve(`playlists/${group.toLowerCase()}.yaml`), contents);
        banner(`Created ${group}.yaml`);
    });
    await Promise.all(promises);
    status.complete('Generating playlists');
}
