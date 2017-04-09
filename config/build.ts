#!/usr/bin/env node --harmony

import * as path from 'path';
import { isEmpty, isString, forIn, isArray, sortBy } from 'lodash';
import * as chalk from 'chalk';
import { status } from './status';
import { SnippetFileInput, SnippetProcessedData, rmRf, mkDir, getFiles, writeFile, banner, loadFileContents } from './helpers';
import { getShareableYaml } from './snippet.helpers';
import { startCase, groupBy, map } from 'lodash';
import { Dictionary } from '@microsoft/office-js-helpers';
import * as jsyaml from 'js-yaml';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/filter';

const { GH_ACCOUNT, GH_REPO, GH_BRANCH } = process.env;
const processedSnippets = new Dictionary<SnippetProcessedData>();
const snippetFilesToUpdate: { [fullPath: string]: string } = {};
const accumulatedErrors: Array<string | Error> = [];


(() => {
    Promise.resolve()
        .then(processSnippets)
        .then(updateModifiedFiles)
        .then(checkSnippetsForUniqueIDs)
        .then(generatePlaylists)
        .then(() => {
            if (accumulatedErrors.length > 0) {
                throw accumulatedErrors;
            }
        })
        .then(() => banner('Done!'))
        .catch(handleError);
})();


async function processSnippets() {
    return new Promise((resolve, reject) => {
        /* Loading samples */
        banner('Loading & processing snippets');
        let files$ = getFiles(path.resolve('samples'), path.resolve('samples'));

        files$.mergeMap(processAndValidateSnippet)
            .filter(file => file !== null)
            .map(file => processedSnippets.add(file.rawUrl, file))
            .subscribe(null, reject, resolve);
    });


    // Helpers:

    async function processAndValidateSnippet(file: SnippetFileInput): Promise<SnippetProcessedData> {
        const messages: Array<string | Error> = [];
        try {
            const localPath = `${file.host}/${file.group}/${file.file_name}`;
            const fullPath = path.resolve('samples', file.path);

            status.add(`Processing ${localPath}`);

            const originalFileContents = await loadFileContents(fullPath);
            let snippet: ISnippet = jsyaml.safeLoad(originalFileContents);


            // Do validations & auto-corrections
            validateStringFieldNotEmptyOrThrow(snippet, 'name');
            validateStringFieldNotEmptyOrThrow(snippet, 'description');
            validateId(snippet, localPath, messages);

            // Additional fields relative to what is normally exposed in sharing
            // (and/or that would normally get erased when doing an export):
            const additionalFields: ISnippet = <any>{};
            additionalFields.id = snippet.id;
            additionalFields.api_set = snippet.api_set;
            additionalFields.author = 'Microsoft';
            if ((typeof (additionalFields as any).order) != 'undefined') {
                // # for ordering, if present (used for samples only)
                (additionalFields as any).order = (snippet as any).order;
            }

            let finalFileContents = getShareableYaml(snippet, additionalFields);
            if (originalFileContents !== finalFileContents) {
                messages.push('Final snippet != original snippet. Queueing to write in new changes.');
                snippetFilesToUpdate[fullPath] = finalFileContents;
            }

            status.complete(true /*success*/, `Processing ${localPath}`, messages);

            const rawUrl = 'https://raw.githubusercontent.com/' +
                `${GH_ACCOUNT || '<ACCOUNT>'}/${GH_REPO || '<REPO>'}/${GH_BRANCH || '<BRANCH>'}` +
                `/samples/${file.host}/${file.group}/${file.file_name}`;

            return {
                id: snippet.id,
                name: snippet.name,
                fileName: file.file_name,
                localPath: localPath,
                description: snippet.description,
                host: file.host,
                rawUrl: rawUrl,
                group: startCase(file.group)
            };

        } catch (exception) {
            messages.push(exception)
            status.complete(false /*success*/, `Processing ${file.host}::${file.file_name}`, messages);
            accumulatedErrors.push(`Failed to process ${file.host}::${file.file_name}: ${exception.message || exception}`);
            return null;
        }
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

    function validateId(snippet: ISnippet, localPath: string, messages: any[]): void {
        // Don't want empty IDs -- or GUID-y IDs either, since they're not particularly memorable...
        if (isEmpty(snippet.id) || isCUID(snippet.id)) {
            snippet.id = localPath.trim().toLowerCase()
                .replace(/[^0-9a-zA-Z]/g, '_') /* replace any non-alphanumeric with an underscore */
                .replace(/_+/g, '_') /* and ensure that don't end up with __ or ___, just a single underscore */
                .replace(/yaml$/i, '') /* remove "yaml" suffix (the ".", now "_", will get removed via underscore-trimming below) */
                .replace(/^_+/, '') /* trim any underscores before */
                .replace(/_+$/, '') /* and trim any at the end, as well */;

            messages.push('Snippet ID may not be empty or be a machine-generated ID.');
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
}

async function updateModifiedFiles() {
    banner('Updating modified files');

    const fileWriteRequests = [];
    forIn(snippetFilesToUpdate, (contents, path) => {
        fileWriteRequests.push(
            Promise.resolve()
                .then(async () => {
                    status.add(`Updating ${path}`);
                    await writeFile(path, contents);
                    status.complete(true /*succeeded*/, `Updating ${path}`);
                })
        );
    })

    await Promise.all(fileWriteRequests);
}

function checkSnippetsForUniqueIDs() {
    banner('Testing every snippet for ID uniqueness');

    let idsAllUnique = true; // assume best, until proven otherwise
    processedSnippets.values()
        .forEach(item => {
            status.add(`Testing ID of snippet ${item.localPath}`);
            const otherMatches = processedSnippets.values().filter(anotherItem => anotherItem !== item && anotherItem.id === item.id);
            const isUnique = (otherMatches.length === 0);
            status.complete(isUnique /*succeeded*/,
                `Testing ID of snippet ${item.localPath}`,
                isUnique ? null : [`ID "${item.id}" not unique, and matches the IDs of `].concat(otherMatches.map(item => '    ' + item.localPath)));
            if (!isUnique) {
                idsAllUnique = false;
            }
        });

    if (!idsAllUnique) {
        throw new Error('Not all snippet IDs are unique; cannot continue');
    }
}

async function generatePlaylists() {
    banner('Generating playlists');

    /* Creating playlists directory */
    status.add('Creating \'playlists\' folder');
    await rmRf('playlists');
    await mkDir('playlists');
    status.complete(true /*success*/, 'Creating \'playlists\' folder');


    const groups = groupBy(
        processedSnippets.values()
            .filter((file) => !(file == null) && file.fileName !== 'default.yaml'),
        'host');
    let playlistPromises = map(groups, async (items, host) => {
        const creatingStatusText = `Creating ${host}.yaml`;
        status.add(creatingStatusText);
        items = sortBy(items, ['group', 'order', 'id']);
        let contents = jsyaml.safeDump(items);
        await writeFile(path.resolve(`playlists/${host}.yaml`), contents);
        status.complete(true /*success*/, creatingStatusText);
    });

    await Promise.all(playlistPromises);
}

function handleError(error: any | any[]) {
    if (!isArray(error)) {
        error = [error];
    }

    banner('One more more errors had occurred during processing:', null, chalk.bold.red);
    (error as any[]).forEach(() => {
        const statusMessage = error.message || error;
        status.add(statusMessage);
        status.complete(false /*successe*/, statusMessage)
    });

    process.exit(1);
}
