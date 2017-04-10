#!/usr/bin/env node --harmony

import * as path from 'path';
import { isNil, isString, isArray, isEmpty, sortBy, cloneDeep } from 'lodash';
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
const snippetFilesToUpdate: Array<{ path: string; contents: string }> = [];
const accumulatedErrors: Array<string | Error> = [];

const officeHosts = ['ACCESS', 'EXCEL', 'ONENOTE', 'OUTLOOK', 'POWERPOINT', 'PROJECT', 'WORD'];
const defaultApiSets = {
    'EXCEL': {
        'ExcelApi': 1.1
    },
    'ONENOTE': {
        'OneNoteApi': 1.1
    },
    'WORD': {
        'WordApi': 1.1
    }

    /* any other hosts is allowed to have no API sets specified*/
};


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
            validateSnippetHost(snippet, file.host, messages);
            validateAtTypesDeclarations(snippet, messages);
            validateOfficialOfficeJs(snippet, file.host, messages);
            validateApiSetNonEmpty(snippet, file.host, localPath, messages);

            // Additional fields relative to what is normally exposed in sharing
            // (and/or that would normally get erased when doing an export):
            const additionalFields: ISnippet = <any>{};
            additionalFields.id = snippet.id;
            additionalFields.api_set = snippet.api_set;
            additionalFields.author = 'Microsoft';
            if ((typeof (additionalFields as any).order) !== 'undefined') {
                // # for ordering, if present (used for samples only)
                (additionalFields as any).order = (snippet as any).order;
            }

            let finalFileContents = getShareableYaml(snippet, additionalFields);
            if (originalFileContents !== finalFileContents) {
                messages.push(chalk.bold.yellow('Final snippet != original snippet. Queueing to write in new changes.'));
                snippetFilesToUpdate.push({ path: fullPath, contents: finalFileContents });
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
                group: startCase(file.group),

                /**
                 * Necessary for back-compat with currently (April 2017)-deployed ScriptLab.
                 * Going forward, though, we want to simply use "rawUrl", as that's more correct semantically.
                 **/
                gist: rawUrl
            };

        } catch (exception) {
            messages.push(exception);
            status.complete(false /*success*/, `Processing ${file.host}::${file.file_name}`, messages);
            accumulatedErrors.push(`Failed to process ${file.host}::${file.file_name}: ${exception.message || exception}`);
            return null;
        }
    }

    function validateStringFieldNotEmptyOrThrow(snippet: ISnippet, field: string): void {
        if (isNil(snippet[field])) {
            throw `Snippet ${field} may not be empty`;
        }

        if (!isString(snippet[field])) {
            throw `Snippet ${field} must be a string`;
        }

        snippet[field] = snippet[field].trim();
    }

    function validateSnippetHost(snippet: ISnippet, host: string, messages: any[]): void {
        host = host.toUpperCase();

        if (typeof snippet.host === 'undefined') {
            messages.push(`Snippet is missing "host" property. Settings based on file path to ${host}`);
            snippet.host = host;
        }

        if (snippet.host !== snippet.host.toUpperCase()) {
            messages.push(`Snippet host is inconsistently-cased. Changing to all-caps`);
            snippet.host = snippet.host.toUpperCase();
        }

        if (snippet.host !== host) {
            throw new Error(`Snippet's specified host "${snippet.host}" is different than the directory path host "${host}". Please fix the mismatch.`);
        }
    }

    function validateAtTypesDeclarations(snippet: ISnippet, messages: any[]) {
        snippet.libraries.split('\n')
            .map(reference => reference.trim())
            .filter(reference => reference.match(/^dt~.*$/gi))
            .map(reference => {
                const atTypesNotation = `@types/${reference.substr('dt~'.length)}`;
                snippet.libraries = snippet.libraries.replace(reference, atTypesNotation);
                messages.push(`Replacing reference "${reference}" with the @types notation: "${atTypesNotation}"`);
            });
    }

    function validateOfficialOfficeJs(snippet: ISnippet, host: string, messages: any[]): void {
        const isOfficeSnippet = officeHosts.indexOf(host.toUpperCase()) >= 0;
        const canonicalOfficeJsReference = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
        const officeDTS = '@types/office-js';

        const officeJsReferences =
            snippet.libraries.split('\n')
                .map(reference => reference.trim())
                .filter(reference => reference.match(/^http.*\/office\.js$/gi));

        const officeJsDTSReference =
            snippet.libraries.split('\n')
                .map(reference => reference.trim())
                .filter(reference => reference === officeDTS);

        if (!isOfficeSnippet) {
            if (officeJsReferences.length > 0 || officeJsDTSReference.length > 0) {
                throw new Error(`Snippet for host "${host}" should not have a reference to Office.js or ${officeDTS}`);
            }
            return;
        }

        // From here on out, can assume that is an Office snippet;
        if (officeJsReferences.length === 0 || officeJsDTSReference.length === 0) {
            throw new Error(`Snippet for host "${host}" should have a reference to Office.js and ${officeDTS}`);
        }

        if (officeJsReferences.length > 1 || officeJsDTSReference.length === 0) {
            throw new Error(`Cannot have more than one reference to Office.js or ${officeDTS}`);
        }

        if (officeJsReferences[0] !== canonicalOfficeJsReference) {
            messages.push(`Office.js reference "${officeJsReferences[0]}" is not in the canonical form of "${canonicalOfficeJsReference}". Fixing it.`);
            snippet.libraries = snippet.libraries.replace(officeJsReferences[0], canonicalOfficeJsReference);
        }
    }

    function validateApiSetNonEmpty(snippet: ISnippet, host: string, localPath: string, messages: any[]): void {
        host = host.toUpperCase();

        if (typeof snippet.api_set === 'undefined') {
            snippet.api_set = {};
        }

        if (isEmpty(snippet.api_set)) {
            if (typeof defaultApiSets[host] === 'undefined') {
                // No API set required (not a host with host-specific APIs), so just exit the function
                return;
            }

            messages.push(new Error(`No API set specified. If building locally, substituting with a default of ` +
                `"${JSON.stringify(defaultApiSets[host])}", but failing the build.`));
            messages.push(new Error('   Please check your pending changes to see the substituted version.'));

            snippet.api_set = cloneDeep(defaultApiSets[host]);

            accumulatedErrors.push(new Error(`No API set specified for ${localPath}`));
        }
    }

    function validateId(snippet: ISnippet, localPath: string, messages: any[]): void {
        // Don't want empty IDs -- or GUID-y IDs either, since they're not particularly memorable...
        if (isNil(snippet.id) || isCUID(snippet.id)) {
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
    const hasChanges = snippetFilesToUpdate.length > 0;
    banner('Updating modified files',
        hasChanges ? null : '<No files to modify>',
        hasChanges ? chalk.bold.yellow : null);

    const fileWriteRequests = [];
    snippetFilesToUpdate.forEach(item => {
        fileWriteRequests.push(
            Promise.resolve()
                .then(async () => {
                    const updatingStatusText = `Updating ${item.path}`;
                    status.add(updatingStatusText);
                    await writeFile(item.path, item.contents);
                    status.complete(true /*succeeded*/,updatingStatusText);
                })
        );
    });

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
    (error as any[]).forEach(item => {
        const statusMessage = item.message || item;
        status.add(statusMessage);
        status.complete(false /*success*/, statusMessage);
    });

    banner('Cannot continue, closing.',
        'Note that if you were building locally, please see pending changes ' + 
        'for anything that may have modified locally (and which might make you pass on a 2nd try).',
        chalk.bold.red);

    process.exit(1);
}
