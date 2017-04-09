#!/usr/bin/env node --harmony

import * as path from 'path';
import { isEmpty, isString } from 'lodash';
import * as chalk from 'chalk';
import { status } from './status';
import { rmRf, mkDir, getFiles, writeFile, banner, loadFileContents } from './helpers';
import { getShareableYaml } from './snippet.helpers';
import { startCase, groupBy, map } from 'lodash';
import { Dictionary } from '@microsoft/office-js-helpers';
import * as jsyaml from 'js-yaml';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/filter';

const { GH_ACCOUNT, GH_REPO, GH_BRANCH } = process.env;
const files = new Dictionary<File>();

(async () => {
    try {
        /* Creating playlists directory */
        status.add('Creating \'playlists\' folder');
        await rmRf('playlists');
        await mkDir('playlists');
        status.complete(true /*success*/, 'Creating \'playlists\' folder');

        /* Loading samples */
        status.add('Loading snippets');
        let files$ = getFiles(path.resolve('samples'), path.resolve('samples'));
        status.complete(true /*success*/, 'Loading snippets');

        files$.mergeMap(async (file) => {
            const messages: Array<string | Error> = [];
            try {
                const hostFilename = `${file.host}::${file.file_name}`;
                const fullPath = path.resolve('samples', file.path);
                status.add(`Processing ${hostFilename}`);
                const originalFileContents = await loadFileContents(fullPath);
                let snippet: ISnippet = jsyaml.safeLoad(originalFileContents);

                // Do validations & auto-corrections
                validateStringFieldNotEmptyOrThrow(snippet, 'name');
                validateStringFieldNotEmptyOrThrow(snippet, 'description');
                validateId(snippet, messages);

                // Additional fields relative to what is normally exposed in a public gist
                // (and/or that would normally get erased when doing an export):
                const additionalFields: ISnippet = <any> {};
                additionalFields.id = snippet.id;
                additionalFields.api_set = snippet.api_set;
                additionalFields.author = 'Microsoft';
                
                messages.push("Before sharing yaml");
                let finalFileContents = getShareableYaml(snippet, additionalFields);
                if (originalFileContents !== finalFileContents) {
                    messages.push("Right before writing...")
                    writeFile(fullPath, finalFileContents);
                    messages.push('Final snippet != original snippet. Writing in the new changes.');
                }

                status.complete(true /*success*/, `Processing ${hostFilename}`, messages);

                return {
                    id: snippet.id,
                    name: snippet.name,
                    fileName: file.file_name,
                    description: snippet.description,
                    host: file.host,
                    gist: `https://raw.githubusercontent.com/${GH_ACCOUNT}/${GH_REPO}/${GH_BRANCH}/samples/${file.host}/${file.group}/${file.file_name}`,
                    group: startCase(file.group)
                };

            } catch (exception) {
                messages.push(exception)
                status.complete(false /*success*/, `Processing ${file.host}::${file.file_name}`, messages);
                handleError(`Failed to process ${file.host}::${file.file_name}: ${exception.message || exception}`);
                return null;
            }
        })
            .filter((file) => !(file == null) && file.fileName !== 'default.yaml')
            .map((file) => files.add(file.gist, file))
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
    if (files.count < 1) {
        return;
    }

    /* Generating playlists */
    status.add('Generating playlists');
    const groups = groupBy(files.values(), 'host');
    let promises = map(groups, async (items, host) => {
        let contents = jsyaml.safeDump(items);
        await writeFile(path.resolve(`playlists/${host}.yaml`), contents);
        banner(`Created ${host}.yaml`);
    });
    await Promise.all(promises);
    status.complete(true /*success*/, 'Generating playlists');
}

function validateStringFieldNotEmptyOrThrow(snippet: ISnippet, field: string): void {
    if (isEmpty(snippet[field])) {
        throw `Snippet ${field} may not be empty`;
    }

    if (!isString(snippet[field])) {
        throw `Snippet ${field} must be a string`;
    }

    snippet[field] = snippet[field].trim();
}

function validateId(snippet: ISnippet, messages: any[]): void {
    // Don't want empty IDs -- or GUID-y IDs either, since they're not particularly memorable...
    if (isEmpty(snippet.id) || isCUID(snippet.id)) {
        snippet.id = snippet.name.trim().toLowerCase()
            .replace(/[^0-9a-zA-Z]/g, '_') /* replace any non-alphanumeric with an underscore */
            .replace(/_+/g, '_') /* and ensure that don't end up with __ or ___, just a single underscore */;

        messages.push('Snippet id may not be empty or be a machine-generated ID.');
        messages.push(`... replacing with an ID based on name: "${snippet.id}"`);
    }

    // Helper:
    function isCUID(id: string) {
        if (id.length === 25 && id.indexOf('_') === -1) {
            // not likely to be a real id, with a name of that precise length and all as one word.
            return true;
        }

        return false;
    }
}