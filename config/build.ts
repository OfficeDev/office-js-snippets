#!/usr/bin/env node --harmony

import * as path from 'path';
import { isNil, isString, isArray, isEmpty, sortBy, cloneDeep } from 'lodash';
import * as chalk from 'chalk';
import { status } from './status';
import {
    SnippetFileInput, SnippetProcessedData,
    getDestinationBranch, followsNamingGuidelines, isCUID,
    rmRf, mkDir, getFiles, writeFile, loadFileContents, banner, getPrintableDetails
} from './helpers';
import { getShareableYaml } from './snippet.helpers';
import { processLibraries } from './libraries.processor';
import { startCase, groupBy, map } from 'lodash';
import { Dictionary } from '@microsoft/office-js-helpers';
import * as jsyaml from 'js-yaml';
import 'rxjs/add/operator/map';
import 'rxjs/add/operator/filter';
import escapeStringRegexp = require('escape-string-regexp');


const { GH_ACCOUNT, GH_REPO, TRAVIS_BRANCH } = process.env;
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
        .then(() => {
            banner('Done!', null, chalk.bold.green);
            process.exit(0);
        })
        .catch(handleError);
})();


async function processSnippets() {
    return new Promise((resolve, reject) => {
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
            status.add(`Processing ${file.relativePath}`);

            const fullPath = path.resolve('samples', file.relativePath);
            const originalFileContents = (await loadFileContents(fullPath)).trim();
            let snippet: ISnippet = jsyaml.safeLoad(originalFileContents);

            // Do validations & auto-corrections
            validateStringFieldNotEmptyOrThrow(snippet, 'name');
            validateStringFieldNotEmptyOrThrow(snippet, 'description');
            validateId(snippet, file.relativePath, messages);
            validateSnippetHost(snippet, file.host, messages);
            validateAtTypesDeclarations(snippet, messages);
            validateOfficialOfficeJs(snippet, file.host, messages);
            validateApiSetNonEmpty(snippet, file.host, file.relativePath, messages);
            validateVersionNumbersOnLibraries(snippet, messages);
            validateTabsInsteadOfSpaces(snippet, messages);
            validateProperFabric(snippet);

            // Additional fields relative to what is normally exposed in sharing
            // (and/or that would normally get erased when doing an export):
            const additionalFields: ISnippet = <any>{};
            additionalFields.id = snippet.id;
            additionalFields.api_set = snippet.api_set;

            additionalFields.author = 'Microsoft';
            if (snippet.author !== 'Microsoft') {
                messages.push(`Replacing "author" field with "Microsoft"`);
            }

            if ((typeof (snippet as any).order) !== 'undefined') {
                // # for ordering, if present (used for samples only)
                (additionalFields as any).order = (snippet as any).order;
            }

            // Finally, some fields simply don't apply, and should be deleted.
            delete snippet.gist;

            let finalFileContents = getShareableYaml(snippet, additionalFields).trim();

            let isDifferent = finalFileContents.replace(/\r\n/g, '\n') !== originalFileContents.replace(/\r\n/g, '\n');
            if (isDifferent) {
                messages.push(chalk.bold.yellow('Final snippet != original snippet. Queueing to write in new changes.'));
                snippetFilesToUpdate.push({ path: fullPath, contents: finalFileContents });
            }

            status.complete(true /*success*/, `Processing ${file.relativePath}`, messages);

            const rawUrl = `https://raw.githubusercontent.com/` +
                `${GH_ACCOUNT || '<ACCOUNT>'}/${GH_REPO || '<REPO>'}/${getDestinationBranch(TRAVIS_BRANCH) || '<BRANCH>'}` +
                `/samples/${file.host}/${file.group}/${file.file_name}`;

            if (messages.findIndex(item => item instanceof Error) >= 0) {
                accumulatedErrors.push(`One or more critical errors on ${file.relativePath}`);
            }

            return {
                id: snippet.id,
                name: snippet.name,
                fileName: file.file_name,
                relativePath: file.relativePath,
                description: snippet.description,
                host: file.host,
                rawUrl: rawUrl,
                group: startCase(file.group),
                order: (typeof (snippet as any).order === 'undefined') ? 100 /* nominally 100 */ : (snippet as any).order,
                api_set: snippet.api_set
            };

        } catch (exception) {
            messages.push(exception);
            status.complete(false /*success*/, `Processing ${file.relativePath}`, messages);
            accumulatedErrors.push(`Failed to process ${file.relativePath}: ${exception.message || exception}`);
            return null;
        }
    }

    function validateProperFabric(snippet: ISnippet): void {
        const libs = snippet.libraries.split('\n').map(reference => reference.trim());
        if (libs.indexOf('office-ui-fabric-js@1.4.0/dist/css/fabric.min.css') >= 0) {
            if (libs.indexOf('office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css') <= 0) {
                throw new Error('Fabric reference is specified, without a reference to a corresponding "fabric.components.min.css". Please add this second Fabric reference as well.');
            }
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
        }
    }

    function validateVersionNumbersOnLibraries(snippet: ISnippet, messages: any[]): void {
        const { scriptReferences, linkReferences } = processLibraries(snippet);
        const allWithoutVersionNumbers = ([].concat(scriptReferences).concat(linkReferences) as string[])
            .filter(item => item.startsWith('https://unpkg.com/'))
            .map(item => item.substr('https://unpkg.com/'.length))
            .filter(item => {
                const containsVersionNumberRegex = /^(@[a-zA-Z_-]+\/)?([a-zA-Z_-]+)@[0-9\.]*.*$/;
                /* Tested with:
                        @microsft/office-js-helpers
                            => wrong

                        @microsft/office-js-helpers/lib.js
                            => wrong

                        @microsft/office-js-helpers@0.6.0
                            => right

                        @microsft/office-js-helpers@0.6.0
                            => right

                        jquery@0.6.0/lib.js
                            => right

                        jquery@0.6.0
                            => right

                        jquery/lib.js
                            => wrong

                        foo-bar/
                            => wrong

                        foo-bar@1.5
                            => right
                */

                return !item.match(containsVersionNumberRegex);
            });

        if (allWithoutVersionNumbers.length === 0) {
            return;
        }

        const defaultSubstitutions = {
            'jquery': 'jquery@3.1.1',
            'office-ui-fabric-js/dist/js/fabric.min.js': 'office-ui-fabric-js@1.4.0/dist/js/fabric.min.js',
            '@microsoft/office-js-helpers/dist/office.helpers.min.js': '@microsoft/office-js-helpers@0.6.0/dist/office.helpers.min.js',
            'core-js/client/core.min.js': 'core-js@2.4.1/client/core.min.js',
            'office-ui-fabric-js/dist/css/fabric.min.css': 'office-ui-fabric-js@1.4.0/dist/css/fabric.min.css',
            'office-ui-fabric-js/dist/css/fabric.components.min.css': 'office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css'
        };

        let hadDefaultSubstitution = false;
        allWithoutVersionNumbers.forEach(item => {
            if (defaultSubstitutions[item]) {
                const regex = new RegExp(`(^\s*)${escapeStringRegexp(item)}(.*)`, 'm');
                snippet.libraries = snippet.libraries.replace(regex, ($0, $1, $2) => $1 + defaultSubstitutions[item] + $2);

                messages.push(new Error(`Missing version number on library ${item}. If building locally, substituting with a default of ` +
                    `"${defaultSubstitutions[item]}", but failing the build.`));
                hadDefaultSubstitution = true;
            }

            messages.push(new Error(`Missing version number on library ${item}. A version # is required for NPM packages.`));
        });

        if (hadDefaultSubstitution) {
            messages.push(new Error('Please check your pending changes to see the default-subtituted substituted library version(s).'));
        }
    }

    function validateId(snippet: ISnippet, localPath: string, messages: any[]): void {
        // Don't want empty IDs -- or GUID-y IDs either, since they're not particularly memorable...
        if (isNil(snippet.id) || snippet.id.trim().length === 0 || isCUID(snippet.id)) {
            snippet.id = localPath.trim().toLowerCase()
                .replace(/[^0-9a-zA-Z]/g, '-') /* replace any non-alphanumeric with a hyphen */
                .replace(/_+/g, '-') /* and ensure that don't end up with -- or --, just a single hyphen */
                .replace(/yaml$/i, '') /* remove "yaml" suffix (the ".", now "-", will get removed via hyphen-trimming below) */
                .replace(/^_+/, '') /* trim any hyphens before */
                .replace(/_+$/, '') /* and trim any at the end, as well */;

            messages.push('Snippet ID may not be empty or be a machine-generated ID.');
            messages.push(`... replacing with an ID based on name: "${snippet.id}"`);
        }

        if (snippet.id.indexOf('_') > 0) {
            snippet.id = snippet.id.replace(/_/g, '-');
            messages.push(`Replacing underscores with hyphens in ID "${snippet.id}"`);
        }

        if (!followsNamingGuidelines(snippet.id)) {
            messages.push(new Error('Snippet ID does not follow naming conventions (only lowercase letters, numbers, and hyphens). Please change it.'));
        }
    }

    function validateTabsInsteadOfSpaces(snippet: ISnippet, messages: any[]): void {
        const codeFields = [snippet.template.content, snippet.script.content, snippet.style.content, snippet.libraries];
        if (codeFields.findIndex(code => code.indexOf('\t') >= 0) >= 0) {
            snippet.template.content = snippet.template.content.replace(/\t/g, '    ');
            snippet.script.content = snippet.script.content.replace(/\t/g, '    ');
            snippet.style.content = snippet.style.content.replace(/\t/g, '    ');
            snippet.libraries = snippet.libraries.replace(/\t/g, '    ');

            messages.push('Snippet had one or more fields (template/script/style/libraries) ' +
                'that contained tabs instead of spaces. Replacing everything with 4 spaces.');
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
                    status.complete(true /*succeeded*/, updatingStatusText);
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
            status.add(`Testing ID of snippet ${item.relativePath}`);
            const otherMatches = processedSnippets.values().filter(anotherItem => anotherItem !== item && anotherItem.id === item.id);
            const isUnique = (otherMatches.length === 0);
            status.complete(isUnique /*succeeded*/,
                `Testing ID of snippet ${item.relativePath}`,
                isUnique ? null : [`ID "${item.id}" not unique, and matches the IDs of `].concat(otherMatches.map(item => '    ' + item.relativePath)));
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

        /*
           Having sorted the items -- which may have included a number in the group name! -- remove the group number if any
           Note that by this time, the group names have already had the dash removed from them,
           and would now have <number><space> if anything at all.

            01 basics
            ==> Contains number, strip it

            basics
            ==> No number, keep as is

            hello-world-123-foo
            ==> No number, keep as is
        */
        const groupNumberRegex = /^(\d+\s)?(\w.*)$/;

        items.forEach(item => {
            item.group = item.group.replace(groupNumberRegex, '$2');

            // Also remove "order", it's no longer needed (the snippets themselves are already in an ordered array in the YAML file)
            delete item.order;
        });

        let contents = jsyaml.safeDump(items, {
            skipInvalid: true /* skip "undefined" (e.g., for "order" on some of the snippets) */
        });

        await writeFile(path.resolve(`playlists/${host}.yaml`), contents);

        status.complete(true /*success*/, creatingStatusText);
    });

    await Promise.all(playlistPromises);
}

function handleError(error: any | any[]) {
    if (!isArray(error)) {
        error = [error];
    }

    banner('One or more errors had occurred during processing:', null, chalk.bold.red);
    (error as any[]).forEach(item => {
        console.log(chalk.red(` * ${item.message || item}`));
        if (item instanceof Error) {
            console.log(getPrintableDetails(item, 4 /*indent*/));
        }
    });

    banner('Cannot continue, closing.',
        'Note that if you were building locally, please see pending changes ' +
        'for anything that may have modified locally (and which might make you pass on a 2nd try).',
        chalk.bold.red);

    process.exit(1);
}
